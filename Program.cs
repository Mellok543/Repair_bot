using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;

var settings = AppSettings.Default;
if (string.IsNullOrWhiteSpace(settings.BotToken) || settings.BotToken.Contains("PASTE_YOUR", StringComparison.OrdinalIgnoreCase))
{
    throw new InvalidOperationException("Откройте AppSettings в Program.cs и укажите реальный BotToken.");
}

var tablesDir = Path.GetFullPath(settings.TablesDirectory);
Directory.CreateDirectory(tablesDir);

var app = new BotApp(
    settings.BotToken,
    Path.Combine(tablesDir, settings.ExcelPath),
    Path.Combine(tablesDir, settings.RepairExcelPath),
    Path.Combine(tablesDir, settings.ConsumablesExcelPath),
    settings.CloserIds);
await app.RunAsync();

sealed record AppSettings(
    string BotToken,
    string TablesDirectory,
    string ExcelPath,
    string RepairExcelPath,
    string ConsumablesExcelPath,
    HashSet<long> CloserIds)
{
    public static AppSettings Default => new(
        BotToken: "7796200129:AAFEfT-KBeqsGzfXBBqvbrH_XuP_XrK3gpU",
        TablesDirectory: "/var/lib/repair_bot/excel",
        ExcelPath: "applications.xlsx",
        RepairExcelPath: "repairs.xlsx",
        ConsumablesExcelPath: "consumables.xlsx",
        CloserIds: [992964625, 222222222]
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
    private readonly HashSet<long> _closerIds;
    private readonly Dictionary<long, SessionState> _sessions = new();
    private int _offset;

    public BotApp(string token, string excelPath, string repairExcelPath, string consumablesExcelPath, HashSet<long> closerIds)
    {
        _token = token;
        _closerIds = closerIds;
        _store = new ApplicationStore(excelPath);
        _repairStore = new RepairStore(repairExcelPath);
        _consumablesStore = new ConsumablesStore(consumablesExcelPath);
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

                    if (text == "Отменить заявку")
                    {
                        if (_sessions.Remove(userId))
                        {
                            await SendMessageAsync(chatId, "Заявка отменена.", Keyboards.MainMenu);
                        }
                        else
                        {
                            await SendMessageAsync(chatId, "Нет активной заявки для отмены.", Keyboards.MainMenu);
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

                    if (text is "/start" or "Меню")
                    {
                        _sessions.Remove(userId);
                        await SendMessageAsync(chatId, "Выберите действие:", Keyboards.MainMenu);
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
                        var activeApplications = _store.GetApplications(ApplicationStore.StatusActive);
                        var activeRepairs = _repairStore.GetRepairs(RepairStore.StatusInProgress);
                        var total = activeApplications.Count + activeRepairs.Count;

                        if (total == 0)
                        {
                            await SendMessageAsync(chatId, "Активных заявок пока нет.", Keyboards.MainMenu);
                        }
                        else
                        {
                            await SendMessageAsync(chatId, $"Активные заявки (всего): {total}", Keyboards.MainMenu);

                            await SendMessageAsync(chatId, $"🛩 Заявки на дроны: {activeApplications.Count}", Keyboards.MainMenu);
                            foreach (var item in activeApplications)
                            {
                                await SendMessageAsync(chatId, item.FormatCard(), Keyboards.MainMenu);
                            }

                            await SendMessageAsync(chatId, $"🛠 Ремонт: {activeRepairs.Count}", Keyboards.MainMenu);
                            foreach (var item in activeRepairs)
                            {
                                await SendMessageAsync(chatId, item.FormatCard(), Keyboards.MainMenu);
                            }
                        }

                        continue;
                    }

                    if (text == "Завершенные заявки")
                    {
                        var completedApplications = _store.GetApplications(ApplicationStore.StatusCompleted);
                        var completedRepairs = _repairStore.GetRepairs(RepairStore.StatusCompleted);
                        var total = completedApplications.Count + completedRepairs.Count;

                        if (total == 0)
                        {
                            await SendMessageAsync(chatId, "Завершённых заявок пока нет.", Keyboards.MainMenu);
                        }
                        else
                        {
                            await SendMessageAsync(chatId, $"Завершённые заявки (всего): {total}", Keyboards.MainMenu);

                            await SendMessageAsync(chatId, $"🛩 Заявки на дроны: {completedApplications.Count}", Keyboards.MainMenu);
                            foreach (var item in completedApplications)
                            {
                                await SendMessageAsync(chatId, item.FormatCard(), Keyboards.MainMenu);
                            }

                            await SendMessageAsync(chatId, $"🛠 Ремонт: {completedRepairs.Count}", Keyboards.MainMenu);
                            foreach (var item in completedRepairs)
                            {
                                await SendMessageAsync(chatId, item.FormatCard(), Keyboards.MainMenu);
                            }
                        }

                        continue;
                    }

                    if (text.Equals("/complete", StringComparison.OrdinalIgnoreCase) || text.StartsWith("/complete ", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!_closerIds.Contains(userId))
                        {
                            await SendMessageAsync(chatId, "У вас нет прав завершать заявки.", Keyboards.MainMenu);
                            continue;
                        }

                        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        if (parts.Length != 2 || !long.TryParse(parts[1], out var appId))
                        {
                            await SendMessageAsync(chatId, "Использование: /complete <id>", Keyboards.MainMenu);
                            continue;
                        }

                        var completed = _store.CompleteApplication(appId);
                        await SendMessageAsync(chatId,
                            completed ? $"Заявка #{appId} завершена." : $"Активная заявка #{appId} не найдена.",
                            Keyboards.MainMenu);
                        continue;
                    }

                    if (text.StartsWith("/complete_repair", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!_closerIds.Contains(userId))
                        {
                            await SendMessageAsync(chatId, "У вас нет прав завершать ремонты.", Keyboards.MainMenu);
                            continue;
                        }

                        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        if (parts.Length != 2 || !long.TryParse(parts[1], out var repairId))
                        {
                            await SendMessageAsync(chatId, "Использование: /complete_repair <id>", Keyboards.MainMenu);
                            continue;
                        }

                        var completed = _repairStore.CompleteRepair(repairId);
                        await SendMessageAsync(chatId,
                            completed ? $"Ремонт #{repairId} завершён." : $"Заявка на ремонт #{repairId} не найдена или уже завершена.",
                            Keyboards.MainMenu);
                        continue;
                    }

                    if (!_sessions.TryGetValue(userId, out var session))
                    {
                        await SendMessageAsync(chatId, "Не понял команду. Нажмите /start", Keyboards.MainMenu);
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
                            await SendMessageAsync(chatId, $"Заявка на ремонт создана!\n\n{repair.FormatCard()}", Keyboards.MainMenu);
                        }
                        else if (session.IsConsumablesRequest)
                        {
                            var consumablesId = _consumablesStore.AddRequest(reporter, session.Data);
                            var consumables = _consumablesStore.GetById(consumablesId);
                            await SendMessageAsync(chatId, $"Заявка на комплектующие создана!\n\n{consumables.FormatCard()}", Keyboards.MainMenu);
                        }
                        else
                        {
                            var appId = _store.AddApplication(reporter, session.Data);
                            var appModel = _store.GetById(appId);
                            await SendMessageAsync(chatId, $"Заявка создана!\n\n{appModel.FormatCard()}", Keyboards.MainMenu);
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

    private static bool IsMenuCommand(string text)
    {
        return text is "/start" or "Меню" or "Оставить заявку" or "Активные заявки" or "Завершенные заявки";
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
                    Step = "consumables_needed";
                    return "Необходимо: (Ручной ввод)";
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
            or "consumables_needed" or "consumables_quantity" or "consumables_note";
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
    public static object MainMenu => Keyboard([["Активные заявки", "Завершенные заявки"], ["Оставить заявку"]]);
    public static object RequestMode => Keyboard([["Обычная заявка", "Ремонт"], ["Комплектующие и расходники"]]);
    public static object PilotType => Keyboard([["КТ", "Оптика", "СТ"]]);
    public static object CancelOnly => Keyboard([["Отменить заявку"]]);
    public static object VideoFrequency => Keyboard([["5.8", "3.4", "3.3"], ["1.5", "1.2"]]);
    public static object ControlFrequency => Keyboard([["2.4", "900", "700"], ["500", "300 кузнец"]]);
    public static object RepairUnit => Keyboard([["КТ", "СТ"], ["Оптика", "Мавики"]]);

    public static object ForStep(string step, SessionState? session = null) => step switch
    {
        "request_mode" => RequestMode,
        "pilot_type" => PilotType,
        "drone_type" => DroneType(session),
        "video_frequency" => VideoFrequency,
        "control_frequency" => ControlFrequency,
        "coil_km" => CoilKmByDrone(session),
        "repair_unit" => RepairUnit,
        "callsign" or "pilot_number" or "rx_firmware" or "regularity_domain" or "bind_phrase" or "quantity"
            or "repair_equipment" or "repair_fault" or "repair_quantity" or "repair_note"
            or "consumables_needed" or "consumables_quantity" or "consumables_note" => CancelOnly,
        _ => MainMenu
    };

    private static object DroneType(SessionState? session)
    {
        var options = session?.DroneOptions() ?? [];
        if (options.Length == 0) return MainMenu;

        var rows = options
            .Chunk(2)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        return Keyboard(rows);
    }

    private static object CoilKmByDrone(SessionState? session)
    {
        var options = session?.CoilOptions() ?? [];
        if (options.Length == 0) return MainMenu;

        return Keyboard([options]);
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
    private readonly string _excelPath;
    private readonly object _sync = new();

    private const string SheetName = "Consumables";
    private static readonly string[] Headers =
    [
        "ID",
        "Дата запроса",
        "Запросил",
        "Необходимо",
        "Количество",
        "Примечание"
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
            ws.Cell(row, 4).Value = payload["consumables_needed"];
            ws.Cell(row, 5).Value = payload["consumables_quantity"];
            ws.Cell(row, 6).Value = payload.GetValueOrDefault("consumables_note", "-");

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
            Needed: ws.Cell(row, 4).GetString(),
            Quantity: ws.Cell(row, 5).GetString(),
            Note: ws.Cell(row, 6).GetString());
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
    string Needed,
    string Quantity,
    string Note)
{
    public string FormatCard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"ID: {Id}");
        sb.AppendLine($"Дата запроса: {RequestDate:dd.MM HH:mm}");
        sb.AppendLine($"Запросил: {RequestedBy}");
        sb.AppendLine($"Необходимо: {Needed}");
        sb.AppendLine($"Количество: {Quantity}");
        sb.AppendLine($"Примечание: {Note}");
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
