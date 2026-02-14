using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;

var settings = AppSettings.Default;
var botToken = Environment.GetEnvironmentVariable("BOT_TOKEN");
if (string.IsNullOrWhiteSpace(botToken))
{
    throw new InvalidOperationException("Укажите BOT_TOKEN в переменных окружения.");
}

var tablesDir = Path.GetFullPath(settings.TablesDirectory);
Directory.CreateDirectory(tablesDir);

var app = new BotApp(
    botToken,
    Path.Combine(tablesDir, settings.ExcelPath),
    Path.Combine(tablesDir, settings.RepairExcelPath),
    Path.Combine(tablesDir, settings.ConsumablesExcelPath),
    Path.Combine(tablesDir, settings.AccessExcelPath),
    settings.CloserIds,
    settings.AccessAdminIds,
    settings.AllowedUserIds,
    settings.NotificationUserIds,
    settings.RecommendationNotificationUserIds);
await app.RunAsync();

sealed record AppSettings(
    string TablesDirectory,
    string ExcelPath,
    string RepairExcelPath,
    string ConsumablesExcelPath,
    string AccessExcelPath,
    HashSet<long> CloserIds,
    HashSet<long> AccessAdminIds,
    HashSet<long> AllowedUserIds,
    HashSet<long> NotificationUserIds,
    HashSet<long> RecommendationNotificationUserIds)
{
    public static AppSettings Default => new(
        TablesDirectory: "/var/lib/repair_bot/excel",
        ExcelPath: "applications.xlsx",
        RepairExcelPath: "repairs.xlsx",
        ConsumablesExcelPath: "consumables.xlsx",
        AccessExcelPath: "access_users.xlsx",
        CloserIds: [992964625, 7302929200, 191974662],
        AccessAdminIds: [992964625],
        AllowedUserIds: [992964625, 7302929200, 191974662],
        NotificationUserIds: [992964625, 7302929200, 191974662],
        RecommendationNotificationUserIds: [992964625]
    );
}

sealed class BotApp
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private const string ApiTemplate = "https://api.telegram.org/bot{0}/{1}";

    private readonly string _token;
    private readonly HttpClient _httpClient = new();
    private readonly ApplicationStore _store;
    private readonly RepairStore _repairStore;
    private readonly ConsumablesStore _consumablesStore;
    private readonly AccessStore _accessStore;
    private readonly HashSet<long> _closerIds;
    private readonly HashSet<long> _accessAdminIds;
    private readonly HashSet<long> _allowedUserIds;
    private readonly HashSet<long> _notificationUserIds;
    private readonly HashSet<long> _recommendationNotificationUserIds;
    private readonly Dictionary<long, SessionState> _sessions = new();
    private int _offset;

    public BotApp(
        string token,
        string excelPath,
        string repairExcelPath,
        string consumablesExcelPath,
        string accessExcelPath,
        HashSet<long> closerIds,
        HashSet<long> accessAdminIds,
        HashSet<long> initialAllowedUserIds,
        HashSet<long> notificationUserIds,
        HashSet<long> recommendationNotificationUserIds)
    {
        _token = token;
        _closerIds = closerIds;
        _accessAdminIds = accessAdminIds;
        _store = new ApplicationStore(excelPath);
        _repairStore = new RepairStore(repairExcelPath);
        _consumablesStore = new ConsumablesStore(consumablesExcelPath);
        _accessStore = new AccessStore(accessExcelPath);
        _notificationUserIds = notificationUserIds;
        _recommendationNotificationUserIds = recommendationNotificationUserIds;

        _accessStore.Bootstrap(initialAllowedUserIds, closerIds, accessAdminIds, notificationUserIds, recommendationNotificationUserIds);

        foreach (var profile in _accessStore.GetUsers())
        {
            if (profile.CanUseBot)
            {
                _allowedUserIds.Add(profile.UserId);
            }

            if (profile.CanComplete)
            {
                _closerIds.Add(profile.UserId);
            }

            if (profile.CanManageAccess)
            {
                _accessAdminIds.Add(profile.UserId);
            }

            if (profile.NotifyRequests)
            {
                _notificationUserIds.Add(profile.UserId);
            }

            if (profile.NotifyRecommendations)
            {
                _recommendationNotificationUserIds.Add(profile.UserId);
            }
        }
    }

    public async Task RunAsync()
    {
        Console.WriteLine("Бот запущен...");

        while (true)
        {
            try
            {
                var updates = await GetUpdatesAsync();
                foreach (var update in updates)
                {
                    _offset = update.UpdateId + 1;

                    var message = update.Message;
                    if (message is null || string.IsNullOrWhiteSpace(message.Text) || message.From is null)
                    {
                        continue;
                    }

                    var text = message.Text.Trim();
                    var chatId = message.Chat.Id;
                    var userId = message.From.Id;
                    var reporter = BuildReporter(message.From);
                    _accessStore.UpdateDisplayName(userId, reporter);
                    var canManageAccess = _accessAdminIds.Contains(userId);
                    var hasAccess = canManageAccess || _allowedUserIds.Contains(userId);

                    if (!hasAccess)
                    {
                        _sessions.Remove(userId);
                        await SendMessageAsync(chatId, "У вас нет доступа к боту.", Keyboards.NoAccess);
                        continue;
                    }

                    var canComplete = _closerIds.Contains(userId);
                    var mainMenu = Keyboards.MainMenu(canComplete, canManageAccess, true);

                    if (text == "Отменить заявку")
                    {
                        if (_sessions.Remove(userId))
                        {
                            await SendMessageAsync(chatId, "Заявка отменена.", mainMenu);
                        }
                        else
                        {
                            await SendMessageAsync(chatId, "Нет активной заявки для отмены.", mainMenu);
                        }

                        continue;
                    }

                    if (_sessions.TryGetValue(userId, out var activeSession) &&
                        activeSession.IsManualInputStep() &&
                        IsMenuCommand(text))
                    {
                        await SendMessageAsync(chatId,
                            "Сейчас идёт ручной ввод. Завершите текущий шаг, чтобы открыть меню.",
                            Keyboards.ForStep(activeSession.Step, activeSession));
                        continue;
                    }

                    if (_sessions.TryGetValue(userId, out var menuSession) &&
                        await TryHandleMenuSessionAsync(menuSession, text, chatId, mainMenu, canManageAccess))
                    {
                        if (menuSession.Step == "done")
                        {
                            _sessions.Remove(userId);
                        }

                        continue;
                    }

                    if (text is "/start" or "Меню")
                    {
                        _sessions.Remove(userId);
                        await SendMessageAsync(chatId, "Выберите действие:", mainMenu);
                        continue;
                    }

                    if (text == "Оставить заявку")
                    {
                        _sessions[userId] = new SessionState("request_mode");
                        await SendMessageAsync(chatId, "Выберите тип заявки:", Keyboards.RequestMode);
                        continue;
                    }

                    if (text == "Активные заявки")
                    {
                        _sessions[userId] = new SessionState("view_active_category");
                        await SendMessageAsync(chatId, "Выберите категорию:", Keyboards.CategoryMenu);
                        continue;
                    }

                    if (text == "Завершенные заявки")
                    {
                        _sessions[userId] = new SessionState("view_completed_category");
                        await SendMessageAsync(chatId, "Выберите категорию:", Keyboards.CategoryMenu);
                        continue;
                    }

                    if (text == "Завершить заявку")
                    {
                        if (!canComplete)
                        {
                            await SendMessageAsync(chatId, "У вас нет прав завершать заявки.", mainMenu);
                            continue;
                        }

                        _sessions[userId] = new SessionState("complete_category");
                        await SendMessageAsync(chatId, "Выберите категорию для завершения:", Keyboards.CategoryMenu);
                        continue;
                    }

                    if (text == "Управление доступом")
                    {
                        if (!canManageAccess)
                        {
                            await SendMessageAsync(chatId, "У вас нет прав управлять доступом.", mainMenu);
                            continue;
                        }

                        _sessions[userId] = new SessionState("access_manage_menu");
                        await SendMessageAsync(chatId, "Управление доступом:", Keyboards.AccessManageMenu);
                        continue;
                    }

                    if (text == "Рекомендовать пользователя")
                    {
                        var recommendationSession = new SessionState("recommend_user_id");
                        recommendationSession.Data["recommender"] = reporter;
                        _sessions[userId] = recommendationSession;
                        await SendMessageAsync(chatId, "Введите Telegram ID пользователя, которого хотите рекомендовать:", Keyboards.CancelOnly);
                        continue;
                    }

                    if (!_sessions.TryGetValue(userId, out var session))
                    {
                        await SendMessageAsync(chatId, "Не понял команду. Нажмите /start", mainMenu);
                        continue;
                    }

                    var nextPrompt = session.Handle(text);
                    if (nextPrompt is null && session.Step == "done")
                    {
                        _sessions.Remove(userId);
                        if (session.IsRepairRequest)
                        {
                            var repairId = _repairStore.AddRepair(reporter, session.Data);
                            var repair = _repairStore.GetById(repairId);
                            await SendMessageAsync(chatId, $"Заявка на ремонт создана!\n\n{repair.FormatCard()}", mainMenu);
                            await NotifyNewRequestAsync($"Новая заявка на ремонт #{repair.Id} от {repair.Reporter}");
                        }
                        else if (session.IsConsumablesRequest)
                        {
                            var consumablesId = _consumablesStore.AddRequest(reporter, session.Data);
                            var consumables = _consumablesStore.GetById(consumablesId);
                            await SendMessageAsync(chatId, $"Заявка на комплектующие создана!\n\n{consumables.FormatCard()}", mainMenu);
                            await NotifyNewRequestAsync($"Новая заявка на комплектующие #{consumables.Id} от {consumables.RequestedBy}");
                        }
                        else
                        {
                            var appId = _store.AddApplication(reporter, session.Data);
                            var appModel = _store.GetById(appId);
                            await SendMessageAsync(chatId, $"Заявка создана!\n\n{appModel.FormatCard()}", mainMenu);
                            await NotifyNewRequestAsync($"Новая заявка на дроны #{appModel.Id} от {appModel.Reporter}");
                        }
                    }
                    else
                    {
                        await SendMessageAsync(chatId, nextPrompt ?? "Продолжайте", Keyboards.ForStep(session.Step, session));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                await Task.Delay(TimeSpan.FromSeconds(2));
            }
        }
    }

    private async Task<bool> TryHandleMenuSessionAsync(SessionState session, string text, long chatId, object mainMenu, bool canManageAccess)
    {
        if (session.Step is "view_active_category" or "view_completed_category")
        {
            var isActive = session.Step == "view_active_category";
            var category = NormalizeCategory(text);
            if (category is null)
            {
                await SendMessageAsync(chatId, "Выберите категорию кнопкой.", Keyboards.CategoryMenu);
                return true;
            }

            if (category == "drones")
            {
                var status = isActive ? ApplicationStore.StatusActive : ApplicationStore.StatusCompleted;
                var items = _store.GetApplications(status);
                await SendCategoryItemsAsync(chatId, isActive ? "Активные" : "Завершённые", "Заявки на дроны", items.Select(i => i.FormatCard()).ToList(), mainMenu);
            }
            else if (category == "repair")
            {
                var status = isActive ? RepairStore.StatusInProgress : RepairStore.StatusCompleted;
                var items = _repairStore.GetRepairs(status);
                await SendCategoryItemsAsync(chatId, isActive ? "Активные" : "Завершённые", "Заявки на ремонт", items.Select(i => i.FormatCard()).ToList(), mainMenu);
            }
            else
            {
                var status = isActive ? ConsumablesStore.StatusInProgress : ConsumablesStore.StatusCompleted;
                var items = _consumablesStore.GetRequests(status);
                await SendCategoryItemsAsync(chatId, isActive ? "Активные" : "Завершённые", "Заявки на комплектующие", items.Select(i => i.FormatCard()).ToList(), mainMenu);
            }

            session.SetStep("done");
            return true;
        }

        if (session.Step == "complete_category")
        {
            var category = NormalizeCategory(text);
            if (category is null)
            {
                await SendMessageAsync(chatId, "Выберите категорию кнопкой.", Keyboards.CategoryMenu);
                return true;
            }

            if (category == "drones")
            {
                var items = _store.GetApplications(ApplicationStore.StatusActive);
                var ids = items.Select(x => x.Id).ToList();
                if (ids.Count == 0)
                {
                    await SendMessageAsync(chatId, "Активных заявок на дроны нет.", mainMenu);
                    session.SetStep("done");
                    return true;
                }

                await SendMessageAsync(chatId, $"Активные заявки на дроны: {ids.Count}", mainMenu);
                foreach (var item in items)
                {
                    await SendMessageAsync(chatId, item.FormatCard(), mainMenu);
                }

                session.SetStep("complete_select_drones");
                await SendMessageAsync(chatId, "Выберите заявку для завершения:", Keyboards.CompleteList(ids));
                return true;
            }

            if (category == "repair")
            {
                var items = _repairStore.GetRepairs(RepairStore.StatusInProgress);
                var ids = items.Select(x => x.Id).ToList();
                if (ids.Count == 0)
                {
                    await SendMessageAsync(chatId, "Активных заявок на ремонт нет.", mainMenu);
                    session.SetStep("done");
                    return true;
                }

                await SendMessageAsync(chatId, $"Активные заявки на ремонт: {ids.Count}", mainMenu);
                foreach (var item in items)
                {
                    await SendMessageAsync(chatId, item.FormatCard(), mainMenu);
                }

                session.SetStep("complete_select_repair");
                await SendMessageAsync(chatId, "Выберите заявку для завершения:", Keyboards.CompleteList(ids));
                return true;
            }

            var consumableItems = _consumablesStore.GetRequests(ConsumablesStore.StatusInProgress);
            var consumableIds = consumableItems.Select(x => x.Id).ToList();
            if (consumableIds.Count == 0)
            {
                await SendMessageAsync(chatId, "Активных заявок на комплектующие нет.", mainMenu);
                session.SetStep("done");
                return true;
            }

            await SendMessageAsync(chatId, $"Активные заявки на комплектующие: {consumableIds.Count}", mainMenu);
            foreach (var item in consumableItems)
            {
                await SendMessageAsync(chatId, item.FormatCard(), mainMenu);
            }

            session.SetStep("complete_select_consumables");
            await SendMessageAsync(chatId, "Выберите заявку для завершения:", Keyboards.CompleteList(consumableIds));
            return true;
        }

        if (session.Step is "complete_select_drones" or "complete_select_repair" or "complete_select_consumables")
        {
            var id = ParseCompleteButton(text);
            if (id is null)
            {
                await SendMessageAsync(chatId, "Выберите заявку кнопкой вида «Завершить #ID».", Keyboards.ForStep(session.Step));
                return true;
            }

            var ok = session.Step switch
            {
                "complete_select_drones" => _store.CompleteApplication(id.Value),
                "complete_select_repair" => _repairStore.CompleteRepair(id.Value),
                _ => _consumablesStore.CompleteRequest(id.Value)
            };

            await SendMessageAsync(chatId, ok ? $"Заявка #{id.Value} завершена." : $"Заявка #{id.Value} не найдена или уже завершена.", mainMenu);
            session.SetStep("done");
            return true;
        }


        if (session.Step == "access_manage_menu")
        {
            if (!canManageAccess)
            {
                await SendMessageAsync(chatId, "У вас нет прав управлять доступом.", mainMenu);
                session.SetStep("done");
                return true;
            }

            if (text == "Добавить пользователя")
            {
                session.SetStep("access_add_user");
                await SendMessageAsync(chatId, "Введите Telegram ID пользователя для выдачи доступа:", Keyboards.CancelOnly);
                return true;
            }

            if (text == "Удалить пользователя")
            {
                session.SetStep("access_remove_user");
                await SendMessageAsync(chatId, "Введите Telegram ID пользователя для удаления доступа:", Keyboards.CancelOnly);
                return true;
            }

            if (text == "Список пользователей")
            {
                var users = _accessStore.GetUsers();
                if (users.Count == 0)
                {
                    await SendMessageAsync(chatId, "Список доступа пуст.", Keyboards.AccessManageMenu);
                    return true;
                }

                await SendMessageAsync(chatId, $"Пользователи: {users.Count}", Keyboards.AccessManageMenu);
                foreach (var user in users)
                {
                    await SendMessageAsync(chatId, user.FormatCard(), Keyboards.AccessManageMenu);
                }

                return true;
            }

            if (text == "Рекомендации")
            {
                var recommendations = _accessStore.GetRecommendations();
                if (recommendations.Count == 0)
                {
                    await SendMessageAsync(chatId, "Рекомендаций пока нет.", Keyboards.AccessManageMenu);
                    return true;
                }

                await SendMessageAsync(chatId, $"Рекомендации: {recommendations.Count}", Keyboards.AccessManageMenu);
                foreach (var recommendation in recommendations)
                {
                    await SendMessageAsync(chatId, recommendation.FormatCard(), Keyboards.AccessManageMenu);
                }

                return true;
            }

            if (text == "Выдать доступ") { session.SetStep("access_grant_bot"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Забрать доступ") { session.SetStep("access_revoke_bot"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Выдать завершение") { session.SetStep("access_grant_complete"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Забрать завершение") { session.SetStep("access_revoke_complete"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Выдать управление") { session.SetStep("access_grant_manage"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Забрать управление") { session.SetStep("access_revoke_manage"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Вкл увед. заявок") { session.SetStep("access_enable_req_notify"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Выкл увед. заявок") { session.SetStep("access_disable_req_notify"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Вкл увед. рек.") { session.SetStep("access_enable_rec_notify"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }
            if (text == "Выкл увед. рек.") { session.SetStep("access_disable_rec_notify"); await SendMessageAsync(chatId, "Введите Telegram ID:", Keyboards.CancelOnly); return true; }

            await SendMessageAsync(chatId, "Выберите действие кнопкой.", Keyboards.AccessManageMenu);
            return true;
        }

        if (session.Step == "access_add_user")
        {
            if (!canManageAccess)
            {
                await SendMessageAsync(chatId, "У вас нет прав управлять доступом.", mainMenu);
                session.SetStep("done");
                return true;
            }

            if (!long.TryParse(text, out var addUserId))
            {
                await SendMessageAsync(chatId, "Введите корректный Telegram ID (число).", Keyboards.CancelOnly);
                return true;
            }

            var added = _accessStore.AddUser(addUserId);
            _allowedUserIds.Add(addUserId);
            await SendMessageAsync(chatId, added ? $"Доступ выдан пользователю {addUserId}." : $"Пользователь {addUserId} уже есть в списке доступа.", mainMenu);
            session.SetStep("done");
            return true;
        }

        if (session.Step == "access_remove_user")
        {
            if (!canManageAccess)
            {
                await SendMessageAsync(chatId, "У вас нет прав управлять доступом.", mainMenu);
                session.SetStep("done");
                return true;
            }

            if (!long.TryParse(text, out var removeUserId))
            {
                await SendMessageAsync(chatId, "Введите корректный Telegram ID (число).", Keyboards.CancelOnly);
                return true;
            }

            if (_accessAdminIds.Contains(removeUserId))
            {
                await SendMessageAsync(chatId, "Нельзя удалить доступ у администратора доступа.", mainMenu);
                session.SetStep("done");
                return true;
            }

            var removed = _accessStore.RemoveUser(removeUserId);
            _allowedUserIds.Remove(removeUserId);
            await SendMessageAsync(chatId, removed ? $"Доступ пользователя {removeUserId} удалён." : $"Пользователь {removeUserId} не найден в списке доступа.", mainMenu);
            session.SetStep("done");
            return true;
        }

        if (session.Step is "access_grant_bot" or "access_revoke_bot" or "access_grant_complete" or "access_revoke_complete" or "access_grant_manage" or "access_revoke_manage" or "access_enable_req_notify" or "access_disable_req_notify" or "access_enable_rec_notify" or "access_disable_rec_notify")
        {
            if (!long.TryParse(text, out var targetUserId))
            {
                await SendMessageAsync(chatId, "Введите корректный Telegram ID (число).", Keyboards.CancelOnly);
                return true;
            }

            switch (session.Step)
            {
                case "access_grant_bot":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: true);
                    _allowedUserIds.Add(targetUserId);
                    await SendMessageAsync(chatId, $"Доступ к боту выдан: {targetUserId}", mainMenu);
                    break;
                case "access_revoke_bot":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: false, canComplete: false, canManageAccess: false, notifyRequests: false, notifyRecommendations: false);
                    _allowedUserIds.Remove(targetUserId);
                    _closerIds.Remove(targetUserId);
                    _accessAdminIds.Remove(targetUserId);
                    _notificationUserIds.Remove(targetUserId);
                    _recommendationNotificationUserIds.Remove(targetUserId);
                    await SendMessageAsync(chatId, $"Доступ к боту снят: {targetUserId}", mainMenu);
                    break;
                case "access_grant_complete":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: true, canComplete: true);
                    _allowedUserIds.Add(targetUserId);
                    _closerIds.Add(targetUserId);
                    await SendMessageAsync(chatId, $"Права завершения выданы: {targetUserId}", mainMenu);
                    break;
                case "access_revoke_complete":
                    _accessStore.UpdatePermissions(targetUserId, canComplete: false);
                    _closerIds.Remove(targetUserId);
                    await SendMessageAsync(chatId, $"Права завершения сняты: {targetUserId}", mainMenu);
                    break;
                case "access_grant_manage":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: true, canManageAccess: true);
                    _allowedUserIds.Add(targetUserId);
                    _accessAdminIds.Add(targetUserId);
                    await SendMessageAsync(chatId, $"Права управления доступом выданы: {targetUserId}", mainMenu);
                    break;
                case "access_revoke_manage":
                    _accessStore.UpdatePermissions(targetUserId, canManageAccess: false);
                    _accessAdminIds.Remove(targetUserId);
                    await SendMessageAsync(chatId, $"Права управления доступом сняты: {targetUserId}", mainMenu);
                    break;
                case "access_enable_req_notify":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: true, notifyRequests: true);
                    _allowedUserIds.Add(targetUserId);
                    _notificationUserIds.Add(targetUserId);
                    await SendMessageAsync(chatId, $"Уведомления о заявках включены: {targetUserId}", mainMenu);
                    break;
                case "access_disable_req_notify":
                    _accessStore.UpdatePermissions(targetUserId, notifyRequests: false);
                    _notificationUserIds.Remove(targetUserId);
                    await SendMessageAsync(chatId, $"Уведомления о заявках выключены: {targetUserId}", mainMenu);
                    break;
                case "access_enable_rec_notify":
                    _accessStore.UpdatePermissions(targetUserId, canUseBot: true, notifyRecommendations: true);
                    _allowedUserIds.Add(targetUserId);
                    _recommendationNotificationUserIds.Add(targetUserId);
                    await SendMessageAsync(chatId, $"Уведомления о рекомендациях включены: {targetUserId}", mainMenu);
                    break;
                default:
                    _accessStore.UpdatePermissions(targetUserId, notifyRecommendations: false);
                    _recommendationNotificationUserIds.Remove(targetUserId);
                    await SendMessageAsync(chatId, $"Уведомления о рекомендациях выключены: {targetUserId}", mainMenu);
                    break;
            }

            session.SetStep("done");
            return true;
        }

        if (session.Step == "recommend_user_id")
        {
            if (!long.TryParse(text, out var recommendedUserId))
            {
                await SendMessageAsync(chatId, "Введите корректный Telegram ID (число).", Keyboards.CancelOnly);
                return true;
            }

            session.Data["recommended_user_id"] = recommendedUserId.ToString();
            session.SetStep("recommend_note");
            await SendMessageAsync(chatId, "Добавьте комментарий (по желанию, отправьте - если без комментария):", Keyboards.CancelOnly);
            return true;
        }

        if (session.Step == "recommend_note")
        {
            var note = string.IsNullOrWhiteSpace(text) || text.Trim() == "-" ? "-" : text.Trim();
            var recommender = session.Data.GetValueOrDefault("recommender", "Неизвестно");
            var recommendedId = long.Parse(session.Data["recommended_user_id"]);
            var recommendationId = _accessStore.AddRecommendation(recommender, recommendedId, note);

            await SendMessageAsync(chatId, $"Рекомендация #{recommendationId} отправлена на рассмотрение.", mainMenu);
            await NotifyRecommendationAsync($"Новая рекомендация #{recommendationId}\nРекомендовал: {recommender}\nКандидат ID: {recommendedId}\nКомментарий: {note}");
            session.SetStep("done");
            return true;
        }

        return false;
    }

    private static string? NormalizeCategory(string text) => text switch
    {
        "Заявки на дроны" => "drones",
        "Заявки на ремонт" => "repair",
        "Заявки на комлектующие" => "consumables",
        "Заявки на комплектующие" => "consumables",
        _ => null
    };

    private static long? ParseCompleteButton(string text)
    {
        if (!text.StartsWith("Завершить #", StringComparison.OrdinalIgnoreCase)) return null;
        var raw = text[11..].Trim();
        return long.TryParse(raw, out var id) ? id : null;
    }

    private async Task SendCategoryItemsAsync(long chatId, string section, string categoryName, List<string> cards, object mainMenu)
    {
        if (cards.Count == 0)
        {
            await SendMessageAsync(chatId, $"{section} {categoryName.ToLower()}: 0", mainMenu);
            return;
        }

        await SendMessageAsync(chatId, $"{section} {categoryName}: {cards.Count}", mainMenu);
        foreach (var card in cards)
        {
            await SendMessageAsync(chatId, card, mainMenu);
        }
    }

    private static bool IsMenuCommand(string text)
    {
        return text is "/start" or "Меню" or "Оставить заявку" or "Активные заявки" or "Завершенные заявки"
            or "Завершить заявку" or "Заявки на дроны" or "Заявки на ремонт" or "Заявки на комлектующие" or "Заявки на комплектующие"
            or "Управление доступом" or "Добавить пользователя" or "Удалить пользователя" or "Список пользователей" or "Рекомендации" or "Рекомендовать пользователя"
            or "Выдать доступ" or "Забрать доступ" or "Выдать завершение" or "Забрать завершение" or "Выдать управление" or "Забрать управление"
            or "Вкл увед. заявок" or "Выкл увед. заявок" or "Вкл увед. рек." or "Выкл увед. рек.";
    }

    private static string BuildReporter(User user)
    {
        if (!string.IsNullOrWhiteSpace(user.Username))
        {
            return $"@{user.Username}";
        }

        var displayName = string.Join(' ', new[] { user.FirstName, user.LastName }
            .Where(x => !string.IsNullOrWhiteSpace(x)));

        return string.IsNullOrWhiteSpace(displayName)
            ? $"tg://user?id={user.Id}"
            : $"{displayName} (tg://user?id={user.Id})";
    }

    private async Task NotifyNewRequestAsync(string message)
    {
        foreach (var notifyUserId in _notificationUserIds)
        {
            try
            {
                await SendMessageAsync(notifyUserId, message, Keyboards.MainMenu(_closerIds.Contains(notifyUserId), _accessAdminIds.Contains(notifyUserId), true));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Не удалось отправить уведомление {notifyUserId}: {ex.Message}");
            }
        }
    }

    private async Task NotifyRecommendationAsync(string message)
    {
        foreach (var adminUserId in _recommendationNotificationUserIds)
        {
            try
            {
                await SendMessageAsync(adminUserId, message, Keyboards.MainMenu(_closerIds.Contains(adminUserId), _accessAdminIds.Contains(adminUserId), true));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Не удалось отправить рекомендацию {adminUserId}: {ex.Message}");
            }
        }
    }

    private async Task<List<Update>> GetUpdatesAsync()
    {
        var url = string.Format(ApiTemplate, _token, "getUpdates") + $"?timeout=25&offset={_offset}";
        using var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();
        var body = await response.Content.ReadAsStringAsync();
        var parsed = JsonSerializer.Deserialize<TgResponse<List<Update>>>(body, JsonOpts);
        return parsed?.Result ?? [];
    }

    private async Task SendMessageAsync(long chatId, string text, object keyboard)
    {
        var url = string.Format(ApiTemplate, _token, "sendMessage");
        var payload = new
        {
            chat_id = chatId,
            text,
            reply_markup = keyboard
        };

        var json = JsonSerializer.Serialize(payload);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await _httpClient.PostAsync(url, content);
        response.EnsureSuccessStatusCode();
    }
}

sealed class SessionState(string step)
{
    private static readonly string[] RepairUnits = ["КТ", "СТ", "Оптика", "Мавики"];
    private static readonly string[] ConsumablesUnits = ["КТ", "СТ", "Мавики"];

    private static readonly Dictionary<string, string[]> DroneTypesByPilotType = new()
    {
        ["КТ"] = ["ПВХ-1", "ПВХ-1Т", "Бумеранг-8", "Бумеранг-10", "Бумеранг-8 День-Ночь", "Бумеранг-10 День-Ночь"],
        ["Оптика"] = ["ПВХ-О", "ПВХ-ОТ", "КВН-День", "КВН День-Ночь"],
        ["СТ"] = ["Молния-1", "Молния-2"]
    };

    private static readonly Dictionary<string, string[]> CoilOptionsByDroneType = new()
    {
        ["ПВХ-О"] = ["15 км", "20 км"],
        ["ПВХ-ОТ"] = ["15 км", "20 км"],
        ["КВН-День"] = ["16 км", "23 км"],
        ["КВН День-Ночь"] = ["16 км", "23 км"]
    };

    public string Step { get; private set; } = step;
    public Dictionary<string, string> Data { get; } = new();
    public bool IsRepairRequest => Data.GetValueOrDefault("request_type") == "repair" || Step.StartsWith("repair_", StringComparison.Ordinal);
    public bool IsConsumablesRequest => Data.GetValueOrDefault("request_type") == "consumables" || Step.StartsWith("consumables_", StringComparison.Ordinal);

    public void SetStep(string stepValue) => Step = stepValue;

    public string? Handle(string text)
    {
        switch (Step)
        {
            case "request_mode":
                if (text == "Обычная заявка")
                {
                    Data["request_type"] = "application";
                    Step = "pilot_type";
                    return "Какой тип?";
                }

                if (text == "Ремонт")
                {
                    Data["request_type"] = "repair";
                    Step = "repair_unit";
                    return "Подразделение:";
                }

                if (text == "Комплектующие и расходники")
                {
                    Data["request_type"] = "consumables";
                    Step = "consumables_unit";
                    return "Подразделение:";
                }

                return "Выберите тип заявки кнопкой: Обычная заявка / Ремонт / Комплектующие и расходники";

            case "pilot_type":
                if (!DroneTypesByPilotType.ContainsKey(text)) return "Выберите тип кнопкой: КТ, Оптика или СТ";
                Data["request_type"] = "application";
                Data["pilot_type"] = text;
                Step = "callsign";
                return "Позывной: (Ручной ввод)";

            case "callsign":
                if (text.Contains("Завершенные заявки")) return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                if (string.IsNullOrWhiteSpace(text)) return "Позывной обязателен. Введите позывной:";
                Data["callsign"] = text.Trim();
                Step = "pilot_number";
                return "Номер пилота. Отправьте '-' если пусто:";

            case "pilot_number":
                if (text == "Завершенные заявки") return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                Data["pilot_number"] = text.Trim() == "-" ? "-" : text.Trim();
                Step = "drone_type";
                return "Тип дрона:";

            case "drone_type":
                if (!Data.TryGetValue("pilot_type", out var type) ||
                    !DroneTypesByPilotType.TryGetValue(type, out var available) ||
                    !available.Contains(text))
                {
                    return "Выберите тип дрона кнопкой.";
                }

                Data["drone_type"] = text;
                if (type == "Оптика")
                {
                    Step = "coil_km";
                    return "Катушка км:";
                }

                Step = "video_frequency";
                return "Частота видео:";

            case "coil_km":
                var availableCoils = CoilOptions();
                if (!availableCoils.Contains(text)) return "Выберите Катушка км кнопкой.";
                Data["coil_km"] = text;
                Data["video_frequency"] = "-";
                Data["control_frequency"] = "-";
                Data["rx_firmware"] = "-";
                Data["regularity_domain"] = "-";
                Data["bind_phrase"] = "-";
                Step = "quantity";
                return "Количество: (Ручной ввод)";

            case "video_frequency":
                if (text is not ("5.8" or "3.4" or "3.3" or "1.5" or "1.2"))
                    return "Выберите частоту видео кнопкой: 5.8 / 3.4 / 3.3 / 1.5 / 1.2";
                Data["video_frequency"] = text;
                Step = "control_frequency";
                return "Частота управления:";

            case "control_frequency":
                if (text is not ("2.4" or "900" or "700" or "500" or "300 кузнец"))
                    return "Выберите частоту управления кнопкой: 2.4 / 900 / 700 / 500 / 300 кузнец";
                Data["control_frequency"] = text;
                Step = "rx_firmware";
                return "Прошивка RX?(Ручной ввод) Пример: Orange5 (beta4)";

            case "rx_firmware":
                if (text == "Завершенные заявки") return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                if (string.IsNullOrWhiteSpace(text)) return "Введите прошивку RX:";
                Data["rx_firmware"] = text.Trim();
                Step = "regularity_domain";
                return "Regularity Domain: (Ручной ввод)";

            case "regularity_domain":
                if (text == "Завершенные заявки") return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                if (string.IsNullOrWhiteSpace(text)) return "Введите Regularity Domain:";
                Data["regularity_domain"] = text.Trim();
                Step = "bind_phrase";
                return "BIND-фраза: (Ручной ввод)";

            case "bind_phrase":
                if (text == "Завершенные заявки") return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                if (string.IsNullOrWhiteSpace(text)) return "BIND-фраза не может быть пустой. Введите значение:";
                Data["bind_phrase"] = text.Trim();
                Step = "quantity";
                return "Количество: (Ручной ввод)";

            case "quantity":
                if (text == "Завершенные заявки") return "Недопустимый выбор";
                if (text == "Активные заявки") return "Недопустимый выбор";
                if (text == "Оставить заявку") return "Недопустимый выбор";
                if (string.IsNullOrWhiteSpace(text)) return "Введите количество:";
                Data["quantity"] = text.Trim();
                if (!Data.ContainsKey("coil_km"))
                {
                    Data["coil_km"] = "-";
                }
                Step = "note";
                return "Примечание: (Ручной ввод, по желанию, отправьте - если пусто)";

            case "note":
                Data["note"] = string.IsNullOrWhiteSpace(text) || text.Trim() == "-" ? "-" : text.Trim();
                Step = "done";
                return null;

            case "repair_unit":
                if (!RepairUnits.Contains(text)) return "Выберите подразделение кнопкой: КТ / СТ / Оптика / Мавики";
                Data["request_type"] = "repair";
                Data["repair_unit"] = text;
                Step = "repair_equipment";
                return "Оборудование: (Ручной ввод)";

            case "repair_equipment":
                if (string.IsNullOrWhiteSpace(text)) return "Введите оборудование:";
                Data["repair_equipment"] = text.Trim();
                Step = "repair_fault";
                return "Неисправность: (Ручной ввод)";

            case "repair_fault":
                if (string.IsNullOrWhiteSpace(text)) return "Введите неисправность:";
                Data["repair_fault"] = text.Trim();
                Step = "repair_quantity";
                return "Количество: (Ручной ввод)";

            case "repair_quantity":
                if (string.IsNullOrWhiteSpace(text)) return "Введите количество:";
                Data["repair_quantity"] = text.Trim();
                Step = "repair_note";
                return "Примечание: (Ручной ввод, по желанию, отправьте - если пусто)";

            case "repair_note":
                Data["repair_note"] = string.IsNullOrWhiteSpace(text) || text.Trim() == "-" ? "-" : text.Trim();
                Step = "done";
                return null;

            case "consumables_unit":
                if (!ConsumablesUnits.Contains(text)) return "Выберите подразделение кнопкой: КТ / СТ / Мавики";
                Data["consumables_unit"] = text;
                Step = "consumables_needed";
                return "Необходимо: (Ручной ввод)";

            case "consumables_needed":
                if (string.IsNullOrWhiteSpace(text)) return "Введите, что необходимо:";
                Data["consumables_needed"] = text.Trim();
                Step = "consumables_quantity";
                return "Количество: (Ручной ввод)";

            case "consumables_quantity":
                if (string.IsNullOrWhiteSpace(text)) return "Введите количество:";
                Data["consumables_quantity"] = text.Trim();
                Step = "consumables_note";
                return "Примечание: (Ручной ввод, по желанию, отправьте - если пусто)";

            case "consumables_note":
                Data["consumables_note"] = string.IsNullOrWhiteSpace(text) || text.Trim() == "-" ? "-" : text.Trim();
                Step = "done";
                return null;

            default:
                return "Ошибка состояния. Нажмите «Оставить заявку» и попробуйте снова.";
        }
    }

    public string[] RepairUnitsOptions() => RepairUnits;

    public string[] CoilOptions()
    {
        if (!Data.TryGetValue("drone_type", out var droneType) ||
            !CoilOptionsByDroneType.TryGetValue(droneType, out var options))
        {
            return [];
        }

        return options;
    }

    public bool IsManualInputStep()
    {
        return Step is "callsign" or "pilot_number" or "rx_firmware" or "regularity_domain" or "bind_phrase" or "quantity" or "note"
            or "repair_equipment" or "repair_fault" or "repair_quantity" or "repair_note"
            or "consumables_needed" or "consumables_quantity" or "consumables_note"
            or "access_add_user" or "access_remove_user"
            or "access_grant_bot" or "access_revoke_bot" or "access_grant_complete" or "access_revoke_complete"
            or "access_grant_manage" or "access_revoke_manage" or "access_enable_req_notify" or "access_disable_req_notify"
            or "access_enable_rec_notify" or "access_disable_rec_notify"
            or "recommend_user_id" or "recommend_note";
    }

    public string[] DroneOptions()
    {
        if (!Data.TryGetValue("pilot_type", out var type) || !DroneTypesByPilotType.TryGetValue(type, out var options))
        {
            return [];
        }

        return options;
    }
}

static class Keyboards
{
    public static object MainMenu(bool canComplete, bool canManageAccess, bool canRecommend)
    {
        var rows = new List<string[]>
        {
            new[] { "Активные заявки", "Завершенные заявки" },
            new[] { "Оставить заявку" }
        };

        if (canComplete)
        {
            rows.Add(new[] { "Завершить заявку" });
        }

        if (canManageAccess)
        {
            rows.Add(new[] { "Управление доступом" });
        }

        if (canRecommend)
        {
            rows.Add(new[] { "Рекомендовать пользователя" });
        }

        return Keyboard(rows.ToArray());
    }

    public static object NoAccess => Keyboard([["/start"]]);
    public static object CategoryMenu => Keyboard([["Заявки на дроны"], ["Заявки на ремонт"], ["Заявки на комлектующие"]]);
    public static object RequestMode => Keyboard([["Обычная заявка", "Ремонт"], ["Комплектующие и расходники"]]);
    public static object AccessManageMenu => Keyboard([["Добавить пользователя", "Удалить пользователя"], ["Выдать доступ", "Забрать доступ"], ["Выдать завершение", "Забрать завершение"], ["Выдать управление", "Забрать управление"], ["Вкл увед. заявок", "Выкл увед. заявок"], ["Вкл увед. рек.", "Выкл увед. рек."], ["Список пользователей", "Рекомендации"], ["Отменить заявку"]]);
    public static object PilotType => Keyboard([["КТ", "Оптика", "СТ"]]);
    public static object CancelOnly => Keyboard([["Отменить заявку"]]);
    public static object VideoFrequency => Keyboard([["5.8", "3.4", "3.3"], ["1.5", "1.2"]]);
    public static object ControlFrequency => Keyboard([["2.4", "900", "700"], ["500", "300 кузнец"]]);
    public static object RepairUnit => Keyboard([["КТ", "СТ"], ["Оптика", "Мавики"]]);
    public static object ConsumablesUnit => Keyboard([["КТ", "СТ", "Мавики"]]);

    public static object ForStep(string step, SessionState? session = null) => step switch
    {
        "request_mode" => RequestMode,
        "pilot_type" => PilotType,
        "drone_type" => DroneType(session),
        "video_frequency" => VideoFrequency,
        "control_frequency" => ControlFrequency,
        "coil_km" => CoilKmByDrone(session),
        "repair_unit" => RepairUnit,
        "consumables_unit" => ConsumablesUnit,
        "callsign" or "pilot_number" or "rx_firmware" or "regularity_domain" or "bind_phrase" or "quantity"
            or "repair_equipment" or "repair_fault" or "repair_quantity" or "repair_note"
            or "consumables_needed" or "consumables_quantity" or "consumables_note"
            or "access_add_user" or "access_remove_user"
            or "access_grant_bot" or "access_revoke_bot" or "access_grant_complete" or "access_revoke_complete"
            or "access_grant_manage" or "access_revoke_manage" or "access_enable_req_notify" or "access_disable_req_notify"
            or "access_enable_rec_notify" or "access_disable_rec_notify"
            or "recommend_user_id" or "recommend_note" => CancelOnly,
        "access_manage_menu" => AccessManageMenu,
        _ => MainMenu(false, false, false)
    };

    private static object DroneType(SessionState? session)
    {
        var options = session?.DroneOptions() ?? [];
        if (options.Length == 0) return MainMenu(false, false, false);

        var rows = options
            .Chunk(2)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        return Keyboard(rows);
    }

    private static object CoilKmByDrone(SessionState? session)
    {
        var options = session?.CoilOptions() ?? [];
        if (options.Length == 0) return MainMenu(false, false, false);

        return Keyboard([options]);
    }

    public static object CompleteList(IEnumerable<long> ids)
    {
        var rows = ids
            .Select(id => $"Завершить #{id}")
            .Chunk(2)
            .Select(chunk => chunk.ToArray())
            .ToList();

        rows.Add(new[] { "Отменить заявку" });
        return Keyboard(rows.ToArray());
    }

    private static object Keyboard(string[][] rows)
    {
        return new
        {
            keyboard = rows.Select(r => r.Select(text => new { text }).ToArray()).ToArray(),
            resize_keyboard = true,
            one_time_keyboard = false
        };
    }
}

sealed class ApplicationStore
{
    public const string StatusActive = "active";
    public const string StatusCompleted = "completed";

    private readonly string _excelPath;
    private readonly object _sync = new();

    private const string SheetName = "Applications";
    private static readonly string[] Headers =
    [
        "ID",
        "Создал",
        "Создано",
        "Завершено",
        "Позывной",
        "Тип",
        "Номер пилота",
        "Тип дрона",
        "Частота видео",
        "Частота управления",
        "Прошивка",
        "Regularity Domain",
        "BIND-фраза",
        "Катушка км",
        "Кол-во",
        "Примечание",
        "Статус"
    ];

    public ApplicationStore(string excelPath)
    {
        _excelPath = excelPath;
        Init();
    }

    public long AddApplication(string reporter, Dictionary<string, string> payload)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);

            var nextId = NextId(ws);
            var row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = nextId;
            ws.Cell(row, 2).Value = reporter;
            ws.Cell(row, 3).Value = DateTime.Now.ToString("s");
            ws.Cell(row, 4).Value = "";
            ws.Cell(row, 5).Value = payload["callsign"];
            ws.Cell(row, 6).Value = payload["pilot_type"];
            ws.Cell(row, 7).Value = string.IsNullOrWhiteSpace(payload.GetValueOrDefault("pilot_number")) ? "-" : payload["pilot_number"];
            ws.Cell(row, 8).Value = payload["drone_type"];
            ws.Cell(row, 9).Value = payload.GetValueOrDefault("video_frequency", "-");
            ws.Cell(row, 10).Value = payload.GetValueOrDefault("control_frequency", "-");
            ws.Cell(row, 11).Value = payload.GetValueOrDefault("rx_firmware", "-");
            ws.Cell(row, 12).Value = payload.GetValueOrDefault("regularity_domain", "-");
            ws.Cell(row, 13).Value = payload.GetValueOrDefault("bind_phrase", "-");
            ws.Cell(row, 14).Value = payload.GetValueOrDefault("coil_km", "-");
            ws.Cell(row, 15).Value = payload["quantity"];
            ws.Cell(row, 16).Value = payload.GetValueOrDefault("note", "-");
            ws.Cell(row, 17).Value = StatusActive;

            workbook.SaveAs(_excelPath);
            return nextId;
        }
    }

    public List<Application> GetApplications(string status)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var result = new List<Application>();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            for (var r = 2; r <= lastRow; r++)
            {
                if (ws.Cell(r, 1).IsEmpty())
                {
                    continue;
                }

                var app = ReadApplication(ws, r);
                if (app.Status == status)
                {
                    result.Add(app);
                }
            }

            return result.OrderByDescending(a => a.Id).ToList();
        }
    }

    public bool CompleteApplication(long appId)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (!long.TryParse(ws.Cell(r, 1).GetString(), out var id) || id != appId)
                {
                    continue;
                }

                var status = ws.Cell(r, 17).GetString();
                if (status != StatusActive)
                {
                    return false;
                }

                ws.Cell(r, 17).Value = StatusCompleted;
                ws.Cell(r, 4).Value = DateTime.Now.ToString("s");
                workbook.SaveAs(_excelPath);
                return true;
            }

            return false;
        }
    }

    public Application GetById(long appId)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id == appId)
                {
                    return ReadApplication(ws, r);
                }
            }

            throw new InvalidOperationException($"Заявка {appId} не найдена");
        }
    }

    private static int FindOrCreateRow(IXLWorksheet ws, long userId)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var existingId) && existingId == userId)
            {
                return r;
            }
        }

        var row = lastRow + 1;
        ws.Cell(row, 1).Value = userId;
        return row;
    }

    private void Init()
    {
        lock (_sync)
        {
            if (!File.Exists(_excelPath))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(SheetName);
                for (var i = 0; i < Headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = Headers[i];
                }

                ws.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
                ws.Columns().AdjustToContents();
                wb.SaveAs(_excelPath);
                return;
            }

            using var existing = new XLWorkbook(_excelPath);
            if (!existing.TryGetWorksheet(SheetName, out var sheet))
            {
                sheet = existing.Worksheets.Add(SheetName);
            }

            for (var i = 0; i < Headers.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = Headers[i];
            }

            sheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
            sheet.Columns().AdjustToContents();
            existing.SaveAs(_excelPath);
        }
    }

    private XLWorkbook OpenWorkbook() => new(_excelPath);

    private static long NextId(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        long max = 0;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    private static Application ReadApplication(IXLWorksheet ws, int row)
    {
        var createdAt = DateTime.Parse(ws.Cell(row, 3).GetString());
        var completedRaw = ws.Cell(row, 4).GetString();
        DateTime? completedAt = string.IsNullOrWhiteSpace(completedRaw) ? null : DateTime.Parse(completedRaw);

        return new Application(
            Id: long.Parse(ws.Cell(row, 1).GetString()),
            Reporter: ws.Cell(row, 2).GetString(),
            CreatedAt: createdAt,
            CompletedAt: completedAt,
            Callsign: ws.Cell(row, 5).GetString(),
            PilotType: ws.Cell(row, 6).GetString(),
            PilotNumber: string.IsNullOrWhiteSpace(ws.Cell(row, 7).GetString()) ? "-" : ws.Cell(row, 7).GetString(),
            DroneType: ws.Cell(row, 8).GetString(),
            VideoFrequency: ws.Cell(row, 9).GetString(),
            ControlFrequency: ws.Cell(row, 10).GetString(),
            RxFirmware: ws.Cell(row, 11).GetString(),
            RegularityDomain: ws.Cell(row, 12).GetString(),
            BindPhrase: ws.Cell(row, 13).GetString(),
            CoilKm: ws.Cell(row, 14).GetString(),
            Quantity: ws.Cell(row, 15).GetString(),
            Note: ws.Cell(row, 16).GetString(),
            Status: ws.Cell(row, 17).GetString());
    }
}

sealed class RepairStore
{
    public const string StatusInProgress = "В работе";
    public const string StatusCompleted = "Завершено";

    private readonly string _excelPath;
    private readonly object _sync = new();

    private const string SheetName = "Repairs";
    private static readonly string[] Headers =
    [
        "ID",
        "Кто передал",
        "Дата передачи",
        "Подразделение",
        "Оборудование",
        "Неисправность",
        "Количество",
        "Примечание",
        "Статус"
    ];

    public RepairStore(string excelPath)
    {
        _excelPath = excelPath;
        Init();
    }

    public long AddRepair(string reporter, Dictionary<string, string> payload)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);

            var nextId = NextId(ws);
            var row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = nextId;
            ws.Cell(row, 2).Value = reporter;
            ws.Cell(row, 3).Value = DateTime.Now.ToString("s");
            ws.Cell(row, 4).Value = payload["repair_unit"];
            ws.Cell(row, 5).Value = payload["repair_equipment"];
            ws.Cell(row, 6).Value = payload["repair_fault"];
            ws.Cell(row, 7).Value = payload["repair_quantity"];
            ws.Cell(row, 8).Value = payload.GetValueOrDefault("repair_note", "-");
            ws.Cell(row, 9).Value = StatusInProgress;

            workbook.SaveAs(_excelPath);
            return nextId;
        }
    }

    public RepairItem GetById(long repairId)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id == repairId)
                {
                    return ReadRepair(ws, r);
                }
            }

            throw new InvalidOperationException($"Ремонт {repairId} не найден");
        }
    }

    public List<RepairItem> GetRepairs(string status)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var result = new List<RepairItem>();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            for (var r = 2; r <= lastRow; r++)
            {
                if (ws.Cell(r, 1).IsEmpty())
                {
                    continue;
                }

                var repair = ReadRepair(ws, r);
                if (repair.Status == status)
                {
                    result.Add(repair);
                }
            }

            return result.OrderByDescending(r => r.Id).ToList();
        }
    }

    public bool CompleteRepair(long repairId)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (!long.TryParse(ws.Cell(r, 1).GetString(), out var id) || id != repairId)
                {
                    continue;
                }

                var status = ws.Cell(r, 9).GetString();
                if (status != StatusInProgress)
                {
                    return false;
                }

                ws.Cell(r, 9).Value = StatusCompleted;
                workbook.SaveAs(_excelPath);
                return true;
            }

            return false;
        }
    }

    private static int FindOrCreateRow(IXLWorksheet ws, long userId)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var existingId) && existingId == userId)
            {
                return r;
            }
        }

        var row = lastRow + 1;
        ws.Cell(row, 1).Value = userId;
        return row;
    }

    private void Init()
    {
        lock (_sync)
        {
            if (!File.Exists(_excelPath))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(SheetName);
                for (var i = 0; i < Headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = Headers[i];
                }

                ws.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
                ws.Columns().AdjustToContents();
                wb.SaveAs(_excelPath);
                return;
            }

            using var existing = new XLWorkbook(_excelPath);
            if (!existing.TryGetWorksheet(SheetName, out var sheet))
            {
                sheet = existing.Worksheets.Add(SheetName);
            }

            for (var i = 0; i < Headers.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = Headers[i];
            }

            sheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
            sheet.Columns().AdjustToContents();
            existing.SaveAs(_excelPath);
        }
    }

    private XLWorkbook OpenWorkbook() => new(_excelPath);

    private static long NextId(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        long max = 0;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    private static RepairItem ReadRepair(IXLWorksheet ws, int row)
    {
        return new RepairItem(
            Id: long.Parse(ws.Cell(row, 1).GetString()),
            Reporter: ws.Cell(row, 2).GetString(),
            TransferDate: DateTime.Parse(ws.Cell(row, 3).GetString()),
            Unit: ws.Cell(row, 4).GetString(),
            Equipment: ws.Cell(row, 5).GetString(),
            Fault: ws.Cell(row, 6).GetString(),
            Quantity: ws.Cell(row, 7).GetString(),
            Note: ws.Cell(row, 8).GetString(),
            Status: ws.Cell(row, 9).GetString());
    }
}


sealed class ConsumablesStore
{
    public const string StatusInProgress = "В работе";
    public const string StatusCompleted = "Завершено";

    private readonly string _excelPath;
    private readonly object _sync = new();

    private const string SheetName = "Consumables";
    private static readonly string[] Headers =
    [
        "ID",
        "Дата запроса",
        "Запросил",
        "Подразделение",
        "Необходимо",
        "Количество",
        "Примечание",
        "Статус"
    ];

    public ConsumablesStore(string excelPath)
    {
        _excelPath = excelPath;
        Init();
    }

    public long AddRequest(string reporter, Dictionary<string, string> payload)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);

            var nextId = NextId(ws);
            var row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = nextId;
            ws.Cell(row, 2).Value = DateTime.Now.ToString("s");
            ws.Cell(row, 3).Value = reporter;
            ws.Cell(row, 4).Value = payload["consumables_unit"];
            ws.Cell(row, 5).Value = payload["consumables_needed"];
            ws.Cell(row, 6).Value = payload["consumables_quantity"];
            ws.Cell(row, 7).Value = payload.GetValueOrDefault("consumables_note", "-");
            ws.Cell(row, 8).Value = StatusInProgress;

            workbook.SaveAs(_excelPath);
            return nextId;
        }
    }

    public ConsumablesItem GetById(long id)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (long.TryParse(ws.Cell(r, 1).GetString(), out var currentId) && currentId == id)
                {
                    return ReadItem(ws, r);
                }
            }

            throw new InvalidOperationException($"Заявка на комплектующие {id} не найдена");
        }
    }

    public List<ConsumablesItem> GetRequests(string status)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var result = new List<ConsumablesItem>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (ws.Cell(r, 1).IsEmpty())
                {
                    continue;
                }

                var item = ReadItem(ws, r);
                if (item.Status == status)
                {
                    result.Add(item);
                }
            }

            return result.OrderByDescending(x => x.Id).ToList();
        }
    }

    public bool CompleteRequest(long id)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (var r = 2; r <= lastRow; r++)
            {
                if (!long.TryParse(ws.Cell(r, 1).GetString(), out var currentId) || currentId != id)
                {
                    continue;
                }

                var status = ws.Cell(r, 8).GetString();
                if (status != StatusInProgress)
                {
                    return false;
                }

                ws.Cell(r, 8).Value = StatusCompleted;
                workbook.SaveAs(_excelPath);
                return true;
            }

            return false;
        }
    }

    private static int FindOrCreateRow(IXLWorksheet ws, long userId)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var existingId) && existingId == userId)
            {
                return r;
            }
        }

        var row = lastRow + 1;
        ws.Cell(row, 1).Value = userId;
        return row;
    }

    private void Init()
    {
        lock (_sync)
        {
            if (!File.Exists(_excelPath))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(SheetName);
                for (var i = 0; i < Headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = Headers[i];
                }

                ws.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
                ws.Columns().AdjustToContents();
                wb.SaveAs(_excelPath);
                return;
            }

            using var existing = new XLWorkbook(_excelPath);
            if (!existing.TryGetWorksheet(SheetName, out var sheet))
            {
                sheet = existing.Worksheets.Add(SheetName);
            }

            for (var i = 0; i < Headers.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = Headers[i];
            }

            sheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
            sheet.Columns().AdjustToContents();
            existing.SaveAs(_excelPath);
        }
    }

    private XLWorkbook OpenWorkbook() => new(_excelPath);

    private static long NextId(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        long max = 0;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    private static ConsumablesItem ReadItem(IXLWorksheet ws, int row)
    {
        return new ConsumablesItem(
            Id: long.Parse(ws.Cell(row, 1).GetString()),
            RequestDate: DateTime.Parse(ws.Cell(row, 2).GetString()),
            RequestedBy: ws.Cell(row, 3).GetString(),
            Unit: ws.Cell(row, 4).GetString(),
            Needed: ws.Cell(row, 5).GetString(),
            Quantity: ws.Cell(row, 6).GetString(),
            Note: ws.Cell(row, 7).GetString(),
            Status: ws.Cell(row, 8).GetString());
    }
}


sealed class AccessStore
{
    private readonly string _excelPath;
    private readonly object _sync = new();

    private const string SheetName = "Users";
    private const string RecommendationSheetName = "Recommendations";
    private static readonly string[] Headers = ["UserId", "DisplayName", "CanUseBot", "CanComplete", "CanManageAccess", "NotifyRequests", "NotifyRecommendations", "AddedAt"];
    private static readonly string[] RecommendationHeaders = ["ID", "Дата", "Рекомендовал", "Кандидат ID", "Комментарий"];

    public AccessStore(string excelPath)
    {
        _excelPath = excelPath;
        Init();
    }

    public List<UserAccessEntry> GetUsers()
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var result = new List<UserAccessEntry>();
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            for (var r = 2; r <= lastRow; r++)
            {
                if (!long.TryParse(ws.Cell(r, 1).GetString(), out var userId))
                {
                    continue;
                }

                result.Add(new UserAccessEntry(
                    UserId: userId,
                    DisplayName: string.IsNullOrWhiteSpace(ws.Cell(r, 2).GetString()) ? "-" : ws.Cell(r, 2).GetString(),
                    CanUseBot: ws.Cell(r, 3).GetString() == "1",
                    CanComplete: ws.Cell(r, 4).GetString() == "1",
                    CanManageAccess: ws.Cell(r, 5).GetString() == "1",
                    NotifyRequests: ws.Cell(r, 6).GetString() == "1",
                    NotifyRecommendations: ws.Cell(r, 7).GetString() == "1"));
            }

            return result.OrderBy(x => x.UserId).ToList();
        }
    }

    public HashSet<long> GetAllUserIds() => GetUsers().Where(x => x.CanUseBot).Select(x => x.UserId).ToHashSet();

    public void Bootstrap(IEnumerable<long> allowed, IEnumerable<long> closers, IEnumerable<long> managers, IEnumerable<long> requestNotify, IEnumerable<long> recommendationNotify)
    {
        foreach (var id in allowed.Distinct())
        {
            UpdatePermissions(id, canUseBot: true);
        }

        foreach (var id in closers.Distinct())
        {
            UpdatePermissions(id, canUseBot: true, canComplete: true);
        }

        foreach (var id in managers.Distinct())
        {
            UpdatePermissions(id, canUseBot: true, canManageAccess: true);
        }

        foreach (var id in requestNotify.Distinct())
        {
            UpdatePermissions(id, canUseBot: true, notifyRequests: true);
        }

        foreach (var id in recommendationNotify.Distinct())
        {
            UpdatePermissions(id, canUseBot: true, notifyRecommendations: true);
        }
    }

    public void UpdateDisplayName(long userId, string displayName)
    {
        if (string.IsNullOrWhiteSpace(displayName))
        {
            return;
        }

        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var row = FindOrCreateRow(ws, userId);
            ws.Cell(row, 2).Value = displayName.Trim();
            workbook.SaveAs(_excelPath);
        }
    }

    public void UpdatePermissions(long userId, bool? canUseBot = null, bool? canComplete = null, bool? canManageAccess = null, bool? notifyRequests = null, bool? notifyRecommendations = null)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(SheetName);
            var row = FindOrCreateRow(ws, userId);

            if (string.IsNullOrWhiteSpace(ws.Cell(row, 2).GetString()))
            {
                ws.Cell(row, 2).Value = "-";
            }

            if (canUseBot.HasValue) ws.Cell(row, 3).Value = canUseBot.Value ? "1" : "0";
            if (canComplete.HasValue) ws.Cell(row, 4).Value = canComplete.Value ? "1" : "0";
            if (canManageAccess.HasValue) ws.Cell(row, 5).Value = canManageAccess.Value ? "1" : "0";
            if (notifyRequests.HasValue) ws.Cell(row, 6).Value = notifyRequests.Value ? "1" : "0";
            if (notifyRecommendations.HasValue) ws.Cell(row, 7).Value = notifyRecommendations.Value ? "1" : "0";
            if (string.IsNullOrWhiteSpace(ws.Cell(row, 8).GetString())) ws.Cell(row, 8).Value = DateTime.Now.ToString("s");

            workbook.SaveAs(_excelPath);
        }
    }

    public bool AddUser(long userId)
    {
        var existed = GetUsers().Any(x => x.UserId == userId);
        UpdatePermissions(userId, canUseBot: true);
        return !existed;
    }

    public bool RemoveUser(long userId)
    {
        var existed = GetUsers().Any(x => x.UserId == userId);
        UpdatePermissions(userId, canUseBot: false, canComplete: false, canManageAccess: false, notifyRequests: false, notifyRecommendations: false);
        return existed;
    }

    public long AddRecommendation(string recommender, long recommendedUserId, string note)
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(RecommendationSheetName);
            var nextId = NextRecommendationId(ws);
            var row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;
            ws.Cell(row, 1).Value = nextId;
            ws.Cell(row, 2).Value = DateTime.Now.ToString("s");
            ws.Cell(row, 3).Value = recommender;
            ws.Cell(row, 4).Value = recommendedUserId;
            ws.Cell(row, 5).Value = note;
            workbook.SaveAs(_excelPath);
            return nextId;
        }
    }

    public List<UserRecommendation> GetRecommendations()
    {
        lock (_sync)
        {
            using var workbook = OpenWorkbook();
            var ws = workbook.Worksheet(RecommendationSheetName);
            var result = new List<UserRecommendation>();

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            for (var r = 2; r <= lastRow; r++)
            {
                if (ws.Cell(r, 1).IsEmpty())
                {
                    continue;
                }

                result.Add(new UserRecommendation(
                    Id: long.Parse(ws.Cell(r, 1).GetString()),
                    CreatedAt: DateTime.Parse(ws.Cell(r, 2).GetString()),
                    Recommender: ws.Cell(r, 3).GetString(),
                    RecommendedUserId: long.Parse(ws.Cell(r, 4).GetString()),
                    Note: ws.Cell(r, 5).GetString()));
            }

            return result.OrderByDescending(x => x.Id).ToList();
        }
    }

    private static long NextRecommendationId(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        long max = 0;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var id) && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    private static int FindOrCreateRow(IXLWorksheet ws, long userId)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            if (long.TryParse(ws.Cell(r, 1).GetString(), out var existingId) && existingId == userId)
            {
                return r;
            }
        }

        var row = lastRow + 1;
        ws.Cell(row, 1).Value = userId;
        return row;
    }

    private void Init()
    {
        lock (_sync)
        {
            if (!File.Exists(_excelPath))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(SheetName);
                for (var i = 0; i < Headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = Headers[i];
                }

                ws.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
                ws.Columns().AdjustToContents();

                var recommendationSheetNew = wb.Worksheets.Add(RecommendationSheetName);
                for (var i = 0; i < RecommendationHeaders.Length; i++)
                {
                    recommendationSheetNew.Cell(1, i + 1).Value = RecommendationHeaders[i];
                }

                recommendationSheetNew.Range(1, 1, 1, RecommendationHeaders.Length).Style.Font.Bold = true;
                recommendationSheetNew.Columns().AdjustToContents();
                wb.SaveAs(_excelPath);
                return;
            }

            using var existing = new XLWorkbook(_excelPath);
            if (!existing.TryGetWorksheet(SheetName, out var sheet))
            {
                sheet = existing.Worksheets.Add(SheetName);
            }

            for (var i = 0; i < Headers.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = Headers[i];
            }

            sheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
            sheet.Columns().AdjustToContents();

            if (!existing.TryGetWorksheet(RecommendationSheetName, out var recommendationSheet))
            {
                recommendationSheet = existing.Worksheets.Add(RecommendationSheetName);
            }

            for (var i = 0; i < RecommendationHeaders.Length; i++)
            {
                recommendationSheet.Cell(1, i + 1).Value = RecommendationHeaders[i];
            }

            recommendationSheet.Range(1, 1, 1, RecommendationHeaders.Length).Style.Font.Bold = true;
            recommendationSheet.Columns().AdjustToContents();
            existing.SaveAs(_excelPath);
        }
    }

    private XLWorkbook OpenWorkbook() => new(_excelPath);
}

record UserAccessEntry(
    long UserId,
    string DisplayName,
    bool CanUseBot,
    bool CanComplete,
    bool CanManageAccess,
    bool NotifyRequests,
    bool NotifyRecommendations)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {UserId}");
        sb.AppendLine($"Юзернейм: {DisplayName}");
        sb.AppendLine($"Доступ к боту: {(CanUseBot ? "Да" : "Нет")}");
        sb.AppendLine($"Завершение заявок: {(CanComplete ? "Да" : "Нет")}");
        sb.AppendLine($"Управление доступом: {(CanManageAccess ? "Да" : "Нет")}");
        sb.AppendLine($"Уведомления о заявках: {(NotifyRequests ? "Да" : "Нет")}");
        sb.AppendLine($"Уведомления о рекомендациях: {(NotifyRecommendations ? "Да" : "Нет")}");
        return sb.ToString().TrimEnd();
    }
}

record UserRecommendation(
    long Id,
    DateTime CreatedAt,
    string Recommender,
    long RecommendedUserId,
    string Note)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {Id}");
        sb.AppendLine($"Дата: {CreatedAt:dd.MM HH:mm}");
        sb.AppendLine($"Рекомендовал: {Recommender}");
        sb.AppendLine($"Кандидат ID: {RecommendedUserId}");
        sb.AppendLine($"Комментарий: {Note}");
        return sb.ToString().TrimEnd();
    }
}

record Application(
    long Id,
    string Reporter,
    DateTime CreatedAt,
    DateTime? CompletedAt,
    string Callsign,
    string PilotType,
    string PilotNumber,
    string DroneType,
    string VideoFrequency,
    string ControlFrequency,
    string RxFirmware,
    string RegularityDomain,
    string BindPhrase,
    string CoilKm,
    string Quantity,
    string Note,
    string Status)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {Id}");
        sb.AppendLine($"Кто оставил: {Reporter}");
        sb.AppendLine($"Дата: {CreatedAt:dd.MM}");
        sb.AppendLine($"Время: {CreatedAt:HH:mm}");
        sb.AppendLine($"Позывной: {Callsign}");
        sb.AppendLine($"Тип: {PilotType}");
        sb.AppendLine($"Номер пилота: {PilotNumber}");
        sb.AppendLine($"Тип дрона: {DroneType}");

        if (PilotType == "Оптика")
        {
            sb.AppendLine($"Катушка км: {CoilKm}");
        }
        else
        {
            sb.AppendLine($"Частота видео: {VideoFrequency}");
            sb.AppendLine($"Частота управления: {ControlFrequency}");
            sb.AppendLine($"Прошивка RX: {RxFirmware}");
            sb.AppendLine($"Regularity Domain: {RegularityDomain}");
            sb.AppendLine($"BIND-фраза: {BindPhrase}");
        }

        sb.AppendLine($"Количество: {Quantity}");
        sb.AppendLine($"Примечание: {Note}");

        if (Status == ApplicationStore.StatusCompleted && CompletedAt.HasValue)
        {
            sb.AppendLine($"Завершено: {CompletedAt:dd.MM HH:mm}");
        }

        return sb.ToString().TrimEnd();
    }
}

record RepairItem(
    long Id,
    string Reporter,
    DateTime TransferDate,
    string Unit,
    string Equipment,
    string Fault,
    string Quantity,
    string Note,
    string Status)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {Id}");
        sb.AppendLine($"Кто передал: {Reporter}");
        sb.AppendLine($"Дата передачи: {TransferDate:dd.MM HH:mm}");
        sb.AppendLine($"Подразделение: {Unit}");
        sb.AppendLine($"Оборудование: {Equipment}");
        sb.AppendLine($"Неисправность: {Fault}");
        sb.AppendLine($"Количество: {Quantity}");
        sb.AppendLine($"Примечание: {Note}");
        sb.AppendLine($"Статус: {Status}");
        return sb.ToString().TrimEnd();
    }
}

record ConsumablesItem(
    long Id,
    DateTime RequestDate,
    string RequestedBy,
    string Unit,
    string Needed,
    string Quantity,
    string Note,
    string Status)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {Id}");
        sb.AppendLine($"Дата запроса: {RequestDate:dd.MM HH:mm}");
        sb.AppendLine($"Запросил: {RequestedBy}");
        sb.AppendLine($"Подразделение: {Unit}");
        sb.AppendLine($"Необходимо: {Needed}");
        sb.AppendLine($"Количество: {Quantity}");
        sb.AppendLine($"Примечание: {Note}");
        sb.AppendLine($"Статус: {Status}");
        return sb.ToString().TrimEnd();
    }
}

record TgResponse<T>
{
    [JsonPropertyName("ok")]
    public bool Ok { get; init; }

    [JsonPropertyName("result")]
    public required T Result { get; init; }
}

record Update
{
    [JsonPropertyName("update_id")]
    public int UpdateId { get; init; }

    [JsonPropertyName("message")]
    public Message? Message { get; init; }
}

record Message
{
    [JsonPropertyName("message_id")]
    public long MessageId { get; init; }

    [JsonPropertyName("chat")]
    public required Chat Chat { get; init; }

    [JsonPropertyName("from")]
    public User? From { get; init; }

    [JsonPropertyName("text")]
    public string? Text { get; init; }
}

record Chat
{
    [JsonPropertyName("id")]
    public long Id { get; init; }
}

record User
{
    [JsonPropertyName("id")]
    public long Id { get; init; }

    [JsonPropertyName("username")]
    public string? Username { get; init; }

    [JsonPropertyName("first_name")]
    public string? FirstName { get; init; }

    [JsonPropertyName("last_name")]
    public string? LastName { get; init; }
}
