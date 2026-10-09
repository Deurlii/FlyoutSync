using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Windows.ApplicationModel.Appointments;

namespace FlyoutSync;

public record SourceConfig(string Name, string Kind, string Url,
    string? Username = null, string? Password = null);

public record Config(List<SourceConfig> Sources);

public record EventItem(string Uid, string Subject, string Location,
    DateTimeOffset Start, DateTimeOffset End, bool AllDay);

public class State
{
    // source name -> LocalIds of the appointments this tool wrote last run
    public Dictionary<string, List<string>> Saved { get; set; } = new();
}

public static class Program
{
    static readonly HttpClient Http = new();
    static readonly DateTimeOffset WindowStart = DateTimeOffset.Now.AddDays(-7);
    static readonly DateTimeOffset WindowEnd = DateTimeOffset.Now.AddDays(60);
    static string ConfigPath => Path.Join(AppContext.BaseDirectory, "flyoutsync.json");
    static string StatePath => Path.Join(AppContext.BaseDirectory, "flyoutsync-state.json");

    public static async Task Main(string[] args)
    {
        var command = args.Length > 0 ? args[0].ToLowerInvariant() : "sync";
        switch (command)
        {
            case "setup":     await Setup(); break;
            case "install":   Install(); break;
            case "uninstall": await Uninstall(); break;
            default:          await RunSync(); break;
        }
        
        var config = JsonSerializer.Deserialize<Config>(
            File.ReadAllText(ConfigPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("Could not read flyoutsync.json. Is it in the same folder as the exe?");

        var state = File.Exists(StatePath)
            ? JsonSerializer.Deserialize<State>(File.ReadAllText(StatePath)) ?? new State()
            : new State();
        
        await File.WriteAllTextAsync(StatePath, JsonSerializer.Serialize(state));

        var store = await AppointmentManager.RequestStoreAsync(
            AppointmentStoreAccessType.AppCalendarsReadWrite);

        foreach (var src in config.Sources)
        {
            try { await SyncSource(store, state, src); }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[{src.Name}] sync failed: {ex.Message}");
            }
        }

        await File.WriteAllTextAsync(StatePath, JsonSerializer.Serialize(state));
    }

    static async Task SyncSource(AppointmentStore store, State state, SourceConfig src)
    {
        var calName = $"Flyout - {src.Name}";

        // Find (or create) this tool's own calendar in the store
        var existing = await store.FindAppointmentCalendarsAsync();
        var calendar = existing.FirstOrDefault(c => c.DisplayName == calName);
        if (calendar is null)
            calendar = await store.CreateAppointmentCalendarAsync(calName);

        // 1. Delete last run's appointments (we tracked their IDs)
        var staleIds = state.Saved.TryGetValue(src.Name, out var ids) ? ids : new List<string>();
        foreach (var id in staleIds)
        {
            try { await calendar.DeleteAppointmentAsync(id); }
            catch { /* already gone */ }
        }

        // 2. Fetch and parse events in the sync window
        var events = src.Kind == "caldav"
            ? await FetchCalDavEvents(src)
            : await FetchIcsEvents(src.Url!);

        // 3. Write them into the appointment store
        var newIds = new List<string>();
        foreach (var ev in events)
        {
            var appt = new Appointment
            {
                Subject = ev.Subject,
                Location = ev.Location,
                StartTime = ev.Start,
                Duration = ev.End - ev.Start,
                AllDay = ev.AllDay,
            };
            await calendar.SaveAppointmentAsync(appt);
            newIds.Add(appt.LocalId);
        }

        state.Saved[src.Name] = newIds;
        Console.WriteLine($"[{src.Name}] {events.Count} events mirrored " +
                          $"({WindowStart:yyyy-MM-dd} to {WindowEnd:yyyy-MM-dd})");
    }

    // ---------- ICS feeds (Google, Outlook) ----------

    static async Task<List<EventItem>> FetchIcsEvents(string url)
    {
        var ics = await Http.GetStringAsync(url);
        var cal = Calendar.Load(ics);
        return ParseCalendar(cal);
    }

    // ---------- iCloud CalDAV ----------

    static readonly XNamespace Dav = "DAV:";
    static readonly XNamespace Cal = "urn:ietf:params:xml:ns:caldav";

    static AuthenticationHeaderValue BasicAuth(SourceConfig s) =>
        new("Basic", Convert.ToBase64String(
            Encoding.UTF8.GetBytes($"{s.Username}:{s.Password}")));

    static async Task<List<EventItem>> FetchCalDavEvents(SourceConfig s)
    {
        // a) discover current-user-principal
        var principal = await Propfind(s.Url!, s,
            "<D:prop xmlns:D='DAV:'><D:current-user-principal/></D:prop>",
            depth: "0",
            pick: x => x.Descendants(Dav + "href").FirstOrDefault()?.Value);
        if (string.IsNullOrEmpty(principal))
            throw new InvalidOperationException("CalDAV: no principal found - check username/password");

        // b) discover calendar-home-set
        var baseUri = new Uri(s.Url!);
        var principalUri = new Uri(baseUri, principal);
        var homeSet = await Propfind(principalUri.ToString(), s,
            "<D:prop xmlns:D='DAV:' xmlns:C='urn:ietf:params:xml:ns:caldav'>" +
            "<C:calendar-home-set/></D:prop>",
            depth: "0",
            pick: x => x.Descendants(Cal + "href").FirstOrDefault()?.Value);
        if (string.IsNullOrEmpty(homeSet))
            throw new InvalidOperationException("CalDAV: no calendar-home-set found");

        // c) REPORT calendar-query: all events overlapping the sync window
        var homeUri = new Uri(principalUri, homeSet).ToString();
        var fmt = "yyyyMMdd'T'HHmmss'Z'";
        var body = new XElement(Cal + "calendar-query",
            new XAttribute(XNamespace.Xmlns + "D", Dav),
            new XAttribute(XNamespace.Xmlns + "C", Cal),
            new XElement(Dav + "prop", new XElement(Dav + "getetag"),
                         new XElement(Cal + "calendar-data")),
            new XElement(Cal + "filter",
                new XElement(Cal + "comp-filter", new XAttribute("name", "VCALENDAR"),
                    new XElement(Cal + "comp-filter", new XAttribute("name", "VEVENT"),
                        new XElement(Cal + "time-range",
                            new XAttribute("start", WindowStart.UtcDateTime.ToString(fmt)),
                            new XAttribute("end", WindowEnd.UtcDateTime.ToString(fmt))))))).ToString();

        using var req = new HttpRequestMessage(new HttpMethod("REPORT"), homeUri);
        req.Headers.Authorization = BasicAuth(s);
        req.Headers.Add("Depth", "1");
        req.Content = new StringContent(body, Encoding.UTF8, "application/xml");

        using var resp = await Http.SendAsync(req);
        resp.EnsureSuccessStatusCode();
        var xml = XDocument.Parse(await resp.Content.ReadAsStringAsync());

        var events = new List<EventItem>();
        foreach (var data in xml.Descendants(Cal + "calendar-data"))
        {
            var ics = (string?)data;
            if (!string.IsNullOrEmpty(ics)) events.AddRange(ParseCalendar(Calendar.Load(ics!)));
        }
        return events;
    }

    static async Task<string?> Propfind(string url, SourceConfig s, string propBody,
        string depth, Func<XDocument, string?> pick)
    {
        using var req = new HttpRequestMessage(new HttpMethod("PROPFIND"), url);
        req.Headers.Authorization = BasicAuth(s);
        req.Headers.Add("Depth", depth);
        req.Content = new StringContent(propBody, Encoding.UTF8, "application/xml");
        using var resp = await Http.SendAsync(req);
        resp.EnsureSuccessStatusCode();
        return pick(XDocument.Parse(await resp.Content.ReadAsStringAsync()));
    }

    // ---------- ICS -> event list (expands recurring events) ----------

    static List<EventItem> ParseCalendar(Ical.Net.Calendar cal)
    {
        var items = new List<EventItem>();
        foreach (var occ in cal.GetOccurrences<CalendarEvent>(
                     WindowStart.UtcDateTime, WindowEnd.UtcDateTime))
        {
            if (occ.Source is not CalendarEvent ev) continue;

            var start = occ.Period.StartTime;
            var end = occ.Period.EndTime ?? occ.Period.StartTime;
            var hasTime = start.HasTime;
            var uid = ev.Uid is null ? Guid.NewGuid().ToString() : ev.Uid;
            var subject = string.IsNullOrEmpty(ev.Summary) ? "(no title)" : ev.Summary!;

            items.Add(new EventItem(
                Uid: uid,
                Subject: subject,
                Location: ev.Location ?? "",
                Start: ToDateTimeOffset(start, hasTime),
                End: ToDateTimeOffset(end, hasTime),
                AllDay: !hasTime));
        }
        return items;
    }
    
    static string ReadPassword() {
        var pw = new System.Text.StringBuilder();
        while (true) {
            var key = Console.ReadKey(intercept: true);
            if (key.Key == ConsoleKey.Enter) { Console.WriteLine(); return pw.ToString(); }
            if (key.Key == ConsoleKey.Backspace && pw.Length > 0) pw.Length--;
            else if (!char.IsControl(key.KeyChar)) pw.Append(key.KeyChar);
        }
    }
    
    static void Install()
    {
        var exe = Environment.ProcessPath!;   // full path of this exe
        var psi = new ProcessStartInfo("schtasks",
            $"/create /tn FlyoutSync /tr \"conhost.exe --headless \"{exe}\"\" /sc minute /mo 15 /f")
        {
            UseShellExecute = false,
            CreateNoWindow = true
        };
        var p = Process.Start(psi)!;
        p.WaitForExit();
        Console.WriteLine(p.ExitCode == 0
            ? "Installed. FlyoutSync will refresh every 15 minutes."
            : "Task registration failed - if your exe path contains spaces, tell me and I'll adjust the quoting.");
    }

    static async Task Uninstall()
    {
        var psi = new ProcessStartInfo("schtasks", "/delete /tn FlyoutSync /f")
        {
            UseShellExecute = false,
            CreateNoWindow = true
        };
        Process.Start(psi)!.WaitForExit();

        var store = await AppointmentManager.RequestStoreAsync(
            AppointmentStoreAccessType.AppCalendarsReadWrite);
        foreach (var cal in await store.FindAppointmentCalendarsAsync())
            if (cal.DisplayName.StartsWith("Flyout - "))
                await cal.DeleteAsync();

        Console.WriteLine("FlyoutSync removed - scheduled task and calendar entries deleted.");
    }
    
    static async Task Setup()
        {
            Console.WriteLine("FlyoutSync setup - mirrors iCloud, Google and Outlook into the taskbar clock.");
            var sources = new List<SourceConfig>();

            if (Ask("Include iCloud?"))
            {
                Console.Write("  Apple ID (e.g. yourname@icloud.com): ");
                var user = Console.ReadLine() ?? "";
                Console.Write("  App-specific password (appleid.apple.com > Sign-In and Security): ");
                sources.Add(new SourceConfig("iCloud", "caldav", "https://caldav.icloud.com/", user, ReadPassword()));
            }

            if (Ask("Include Google?"))
            {
                Console.Write("  Secret iCal address (Google Calendar web > gear > Settings > Integrate calendar): ");
                sources.Add(new SourceConfig("Google", "ics", Console.ReadLine() ?? ""));
            }

            if (Ask("Include Outlook?"))
            {
                Console.Write("  Published ICS link (Outlook web > Settings > Calendar > Shared calendars): ");
                sources.Add(new SourceConfig("Outlook", "ics", Console.ReadLine() ?? ""));
            }

            // Live test sync before we save anything
            var state = new State();
            var store = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            var failures = 0;
            foreach (var src in sources)
            {
                try { await SyncSource(store, state, src); }
                catch (Exception ex) { failures++; Console.Error.WriteLine($"  [{src.Name}] failed: {ex.Message}"); }
            }
            await File.WriteAllTextAsync(StatePath, JsonSerializer.Serialize(state));

            if (failures > 0 && !Ask("At least one source failed. Save this configuration anyway?"))
                return;

            await File.WriteAllTextAsync(ConfigPath,
                JsonSerializer.Serialize(new Config(sources), new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine("Configuration saved next to the exe.");

            if (Ask("Install the 15-minute auto-refresh now?"))
                Install();
    }

    static bool Ask(string question)
    {
        Console.Write($"{question} (y/n): ");
        return (Console.ReadLine() ?? "").Trim().StartsWith("y", StringComparison.OrdinalIgnoreCase);
    }
    
    static async Task RunSync()
    {
        var config = JsonSerializer.Deserialize<Config>(
                         File.ReadAllText(ConfigPath),
                         new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                     ?? throw new InvalidOperationException("Could not read flyoutsync.json.");

        var state = File.Exists(StatePath)
            ? JsonSerializer.Deserialize<State>(File.ReadAllText(StatePath)) ?? new State()
            : new State();

        var store = await AppointmentManager.RequestStoreAsync(
            AppointmentStoreAccessType.AppCalendarsReadWrite);

        foreach (var src in config.Sources)
        {
            try { await SyncSource(store, state, src); }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[{src.Name}] sync failed: {ex.Message}");
            }
        }

        await File.WriteAllTextAsync(StatePath, JsonSerializer.Serialize(state));
    }
    
    static DateTimeOffset ToDateTimeOffset(IDateTime dt, bool hasTime)
    {
        var when = dt.Value;
        if (!hasTime)
        {
            // All-day: anchor at midnight local time
            return new DateTimeOffset(when.Date, TimeZoneInfo.Local.GetUtcOffset(when.Date));
        }
        if (when.Kind == DateTimeKind.Utc || string.IsNullOrEmpty(dt.TzId))
            return new DateTimeOffset(when, TimeSpan.Zero);

        try
        {
            var tz = TimeZoneInfo.FindSystemTimeZoneById(dt.TzId!);
            return new DateTimeOffset(when, tz.GetUtcOffset(when));
        }
        catch (Exception)
        {
            return new DateTimeOffset(when, TimeSpan.Zero);
        }
    }
}
