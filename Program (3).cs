using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;

internal class Program
{
    // =========================================================
    // CONFIG
    // =========================================================

    private const string GmailAddress =
        "samzera22@gmail.com";

    private const string GmailAppPassword =
        "pnhnmkzvcqbysazx";

    private const string MacroDroidWebhookUrl =
        "https://trigger.macrodroid.com/8c5b05e5-f9c7-4e3b-9eea-3b1d0cfc659f/kabo";

    // Exact-match senders (real people - address won't suddenly change).
    private static readonly string[] ExactSenderMatches =
    {
        "kinginthenorth22@hotmail.com"
    };

    // Substring matches on the address (case-insensitive) - covers a
    // business/notification sender even if the exact local-part or
    // domain suffix shifts (info@, noreply@, kabasakal.com,
    // kabasakalonline.net, etc all contain "kabasakal").
    private static readonly string[] PartialSenderMatches =
    {
        "kabasakal"
    };

    // Bu ifadelerden HERHANGİ BİRİ subject/body'de geçerse webhook tetiklenir.
    private static readonly string[] TriggerTexts =
    {
        "kabasakal"
    };

    // true  -> mail ayrıca yukarıdaki gönderen filtresinden de geçmeli
    // false -> gönderen fark etmez, subject/body'de "kabasakal" geçmesi yeter
    private static readonly bool RequireSenderMatch = true;

    // Once a match is found, fire the webhook this many times,
    // waiting this long between each attempt.
    private const int WebhookTriggerCount = 3;

    private static readonly TimeSpan WebhookTriggerDelay =
        TimeSpan.FromSeconds(10);

    private static readonly HttpClient Http =
        new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(15)
        };

    // How long to keep each IDLE command open before renewing it.
    // Kept short (well under most home routers' NAT idle-timeout,
    // typically ~5 minutes) so a silently-dropped connection is
    // detected and reconnected quickly instead of sitting dead.
    private static readonly TimeSpan IdleRenewInterval =
        TimeSpan.FromMinutes(3);

    // Socket-level read/write timeout. Lower than the MailKit
    // default (2 minutes) so a dead socket fails fast instead of
    // hanging silently.
    private static readonly TimeSpan SocketTimeout =
        TimeSpan.FromSeconds(30);

    // =========================================================
    // MAIN
    // =========================================================

    private static Mutex? _singleInstanceMutex;

    public static async Task Main()
    {
        // Prevent accidentally running two copies at once (e.g. from
        // forgetting to close a previous window) - each copy opens 2
        // IMAP connections, and Gmail caps simultaneous IMAP sessions
        // per account, so duplicates can silently starve this one out
        // after a while.
        _singleInstanceMutex =
            new Mutex(true, "GmailMacroDroidMonitor_SingleInstance", out bool isNewInstance);

        if (!isNewInstance)
        {
            Console.WriteLine("==========================================");
            Console.WriteLine("[FATAL] Another copy of this monitor is already running.");
            Console.WriteLine("Check Task Manager for a leftover dotnet/exe process,");
            Console.WriteLine("close it, then run this again.");
            Console.WriteLine("==========================================");
            return;
        }

        Console.Title = "Gmail -> MacroDroid Monitor";

        Console.WriteLine("==========================================");
        Console.WriteLine("       Gmail -> MacroDroid Monitor");
        Console.WriteLine("==========================================");
        Console.WriteLine();
        Console.WriteLine($"Gmail:    {GmailAddress}");
        Console.WriteLine(
            RequireSenderMatch
                ? $"Senders:  {string.Join(", ", ExactSenderMatches)} (exact), " +
                  $"*{string.Join("*, *", PartialSenderMatches)}* (contains)"
                : "Senders:  (any sender)");
        Console.WriteLine(
            $"Required: any of {string.Join(" / ", TriggerTexts.Select(t => $"\"{t}\""))}");
        Console.WriteLine("Mode:     IMAP IDLE (push), watching Inbox + Spam");
        Console.WriteLine();
        Console.WriteLine("Existing emails will be ignored.");
        Console.WriteLine("Waiting for NEW emails...");
        Console.WriteLine("Press Ctrl+C to stop.");
        Console.WriteLine();

        // Run one independent monitor per folder so a message that
        // lands in Spam (very common for automated contact-form mail)
        // is still caught even though it never touches the Inbox.
        var inboxTask =
            MonitorFolderForever(null, "INBOX");

        var spamTask =
            MonitorFolderForever(SpecialFolder.Junk, "SPAM");

        await Task.WhenAll(inboxTask, spamTask);
    }

    // =========================================================
    // PER-FOLDER RECONNECT LOOP
    // =========================================================

    private static async Task MonitorFolderForever(
        SpecialFolder? folderKind,
        string label)
    {
        while (true)
        {
            try
            {
                await MonitorGmail(folderKind, label);
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine($"[{label}] [ERROR] {ex.GetType().Name}: {ex.Message}");

                if (ex.InnerException != null)
                {
                    Console.WriteLine(
                        $"[{label}] [ERROR] Inner: " +
                        $"{ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
                }

                Console.WriteLine($"[{label}] [INFO] Reconnecting in 5 seconds...");

                await Task.Delay(TimeSpan.FromSeconds(5));
            }
        }
    }

    // =========================================================
    // GMAIL MONITOR (single folder, single connection)
    // =========================================================

    private static async Task MonitorGmail(
        SpecialFolder? folderKind,
        string label)
    {
        using var client = new ImapClient
        {
            Timeout = (int)SocketTimeout.TotalMilliseconds
        };

        Console.WriteLine($"[{label}] Connecting...");

        await client.ConnectAsync(
            "imap.gmail.com",
            993,
            SecureSocketOptions.SslOnConnect);

        Console.WriteLine($"[{label}] Authenticating...");

        await client.AuthenticateAsync(
            GmailAddress,
            GmailAppPassword);

        Console.WriteLine($"[{label}] Connected.");

        IMailFolder? folder;

        if (folderKind == null)
        {
            // Inbox is not a "special-use" folder in IMAP terms -
            // it's always just client.Inbox.
            folder = client.Inbox;
        }
        else
        {
            try
            {
                folder = client.GetFolder(folderKind.Value);
            }
            catch (Exception)
            {
                folder = null;
            }
        }

        if (folder == null)
        {
            Console.WriteLine(
                $"[{label}] This special folder does not exist on the " +
                "account (or Gmail hasn't exposed it over IMAP). Skipping.");

            // Don't spin-loop forever reconnecting for a folder
            // that will never exist.
            await Task.Delay(Timeout.Infinite);

            return;
        }

        await folder.OpenAsync(FolderAccess.ReadOnly);

        Console.WriteLine(
            $"[{label}] Opened '{folder.FullName}' " +
            $"({folder.Count} messages currently in it).");

        // =====================================================
        // BASELINE - ignore everything already present
        // =====================================================

        var initialUids = await folder.SearchAsync(SearchQuery.All);

        uint lastSeenUid = 0;

        if (initialUids.Count > 0)
        {
            lastSeenUid = initialUids.Max(x => x.Id);
        }

        Console.WriteLine($"[{label}] Starting UID: {lastSeenUid}");
        Console.WriteLine($"[{label}] [READY] Waiting for NEW emails...");
        Console.WriteLine();

        bool supportsIdle =
            client.Capabilities.HasFlag(ImapCapabilities.Idle);

        if (!supportsIdle)
        {
            Console.WriteLine(
                $"[{label}] [WARN] Server does not support IDLE, " +
                "falling back to polling every 15s.");
        }

        // =====================================================
        // WATCH LOOP
        // =====================================================

        while (client.IsConnected)
        {
            if (supportsIdle)
            {
                var changeSignal = new CancellationTokenSource();

                void OnCountChanged(object? sender, EventArgs e) =>
                    changeSignal.Cancel();

                folder.CountChanged += OnCountChanged;

                try
                {
                    using var renewTimeout =
                        new CancellationTokenSource(IdleRenewInterval);

                    using var linked =
                        CancellationTokenSource.CreateLinkedTokenSource(
                            changeSignal.Token,
                            renewTimeout.Token);

                    // Blocks until either the mailbox changes or the
                    // renew timeout hits - whichever comes first.
                    await client.IdleAsync(linked.Token);
                }
                catch (Exception ex)
                {
                    // If the socket died silently (e.g. router/NAT
                    // dropped an idle connection), this is where it
                    // surfaces. Rethrow so the outer loop reconnects.
                    Console.WriteLine(
                        $"[{label}] [IDLE ERROR] {ex.GetType().Name}: {ex.Message} - reconnecting...");
                    throw;
                }
                finally
                {
                    folder.CountChanged -= OnCountChanged;
                }

                Console.WriteLine(
                    $"[{label}] [HEARTBEAT] {DateTime.Now:HH:mm:ss} - still alive, checking for new mail...");
            }
            else
            {
                await Task.Delay(TimeSpan.FromSeconds(15));
            }

            // Whether we woke up because of IDLE or a poll timer,
            // check what's actually new.
            List<UniqueId> newUids;

            try
            {
                var currentUids = await folder.SearchAsync(SearchQuery.All);

                newUids = currentUids
                    .Where(x => x.Id > lastSeenUid)
                    .OrderBy(x => x.Id)
                    .ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{label}] Check error: {ex.Message}");
                break; // let the outer loop reconnect
            }

            foreach (var uid in newUids)
            {
                lastSeenUid = uid.Id;

                // Hard timeout: whatever happens inside processing a
                // single email (webhook calls x3 with delays, etc), it
                // can never be allowed to freeze this loop forever.
                var processTask = ProcessNewEmail(folder, uid, label);
                var timeoutTask = Task.Delay(TimeSpan.FromSeconds(60));

                var finished = await Task.WhenAny(processTask, timeoutTask);

                if (finished == timeoutTask)
                {
                    Console.WriteLine(
                        $"[{label}] [WARN] Processing UID {uid.Id} took " +
                        "too long, moving on (it will finish in the background).");
                }
            }
        }

        Console.WriteLine($"[{label}] Connection lost.");
    }

    // =========================================================
    // PROCESS NEW EMAIL
    // =========================================================

    private static async Task ProcessNewEmail(
        IMailFolder folder,
        UniqueId uid,
        string label)
    {
        var message = await folder.GetMessageAsync(uid);

        string fromAddress =
            message.From.Mailboxes.FirstOrDefault()?.Address ?? "";

        var replyToAddresses =
            message.ReplyTo.Mailboxes
                .Select(m => m.Address)
                .ToList();

        Console.WriteLine();
        Console.WriteLine($"[{label}] [{DateTime.Now:HH:mm:ss}] NEW EMAIL");
        Console.WriteLine($"    From:     {fromAddress}");

        if (replyToAddresses.Count > 0)
        {
            Console.WriteLine(
                $"    Reply-To: {string.Join(", ", replyToAddresses)}");
        }

        Console.WriteLine($"    Subject:  {message.Subject}");

        // =====================================================
        // CHECK SENDER (From OR Reply-To match)
        // =====================================================

        bool senderMatches =
            IsMatchingSender(fromAddress) ||
            replyToAddresses.Any(IsMatchingSender);

        if (RequireSenderMatch && !senderMatches)
        {
            Console.WriteLine("    -> Sender does not match. Ignored.");
            return;
        }

        // =====================================================
        // CHECK SUBJECT + BODY (any of the trigger texts)
        // =====================================================

        string content = NormalizeForMatch(BuildSearchableContent(message));

        string? matchedText = TriggerTexts.FirstOrDefault(t =>
            content.Contains(NormalizeForMatch(t)));

        if (matchedText == null)
        {
            Console.WriteLine(
                "    -> No trigger text found in subject/body. Ignored.");
            return;
        }

        // =====================================================
        // MATCH
        // =====================================================

        Console.WriteLine();
        Console.WriteLine("==========================================");
        Console.WriteLine("              MATCH FOUND!");
        Console.WriteLine("==========================================");
        Console.WriteLine($"From:    {fromAddress}");
        Console.WriteLine($"Subject: {message.Subject}");
        Console.WriteLine($"Text:    {matchedText}");
        Console.WriteLine();

        bool webhookSuccess =
            await TriggerMacroDroidRepeatedly(message, label, matchedText);

        Console.WriteLine(
            webhookSuccess
                ? "[OK] All MacroDroid webhook triggers completed."
                : "[WARNING] One or more MacroDroid webhook triggers failed.");

        Console.WriteLine();
    }

    // =========================================================
    // SENDER MATCHING
    // =========================================================

    private static bool IsMatchingSender(string address)
    {
        if (string.IsNullOrEmpty(address))
        {
            return false;
        }

        bool exactHit = ExactSenderMatches.Any(t =>
            address.Equals(t, StringComparison.OrdinalIgnoreCase));

        if (exactHit)
        {
            return true;
        }

        bool partialHit = PartialSenderMatches.Any(t =>
            address.Contains(t, StringComparison.OrdinalIgnoreCase));

        return partialHit;
    }

    // =========================================================
    // FIRE THE WEBHOOK MULTIPLE TIMES
    // =========================================================

    private static async Task<bool> TriggerMacroDroidRepeatedly(
        MimeMessage message,
        string label,
        string matchedText)
    {
        bool allSucceeded = true;

        for (int attempt = 1; attempt <= WebhookTriggerCount; attempt++)
        {
            Console.WriteLine(
                $"[{label}] [WEBHOOK] Trigger {attempt}/{WebhookTriggerCount}...");

            bool success = await TriggerMacroDroid(message, matchedText);

            allSucceeded &= success;

            Console.WriteLine(
                success
                    ? $"[{label}] [OK] Trigger {attempt}/{WebhookTriggerCount} succeeded."
                    : $"[{label}] [WARNING] Trigger {attempt}/{WebhookTriggerCount} failed.");

            bool isLastAttempt = attempt == WebhookTriggerCount;

            if (!isLastAttempt)
            {
                await Task.Delay(WebhookTriggerDelay);
            }
        }

        return allSucceeded;
    }

    // =========================================================
    // BUILD SEARCHABLE CONTENT
    // =========================================================
    // Decodes HTML entities and strips tags so a phrase can't be
    // split apart by markup (e.g. "mesaj<br>gönderdi") or hidden
    // behind an &nbsp; instead of a normal space.
    // =========================================================

    private static string BuildSearchableContent(MimeMessage message)
    {
        var sb = new StringBuilder();

        sb.AppendLine(message.Subject ?? "");
        sb.AppendLine(message.TextBody ?? "");
        sb.AppendLine(PlainTextFromHtml(message.HtmlBody ?? ""));

        return sb.ToString();
    }

    private static string PlainTextFromHtml(string html)
    {
        if (string.IsNullOrEmpty(html))
        {
            return "";
        }

        string decoded = WebUtility.HtmlDecode(html);

        // Turn tags/line-breaks into spaces rather than deleting them,
        // so words on either side of a tag don't get glued together.
        string noTags = Regex.Replace(decoded, "<[^>]+>", " ");

        string normalized = noTags.Replace('\u00A0', ' ');

        return normalized;
    }

    // =========================================================
    // NORMALIZE FOR MATCHING
    // =========================================================
    // Türkçe i/İ/ı farklarını ve satır sonu/çoklu boşlukları eşitler.
    // OrdinalIgnoreCase "İ" ile "i" (ve "ı" ile "I") harflerini aynı
    // saymaz; bu yüzden hem mail içeriği hem de aranan ifadeler aynı
    // şekilde normalize edilip karşılaştırılır.
    // =========================================================

    private static string NormalizeForMatch(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return "";
        }

        string folded = text
            .Replace('İ', 'i')
            .Replace('ı', 'i')
            .ToLowerInvariant();

        return Regex.Replace(folded, @"\s+", " ").Trim();
    }

    // =========================================================
    // MACRODROID WEBHOOK
    // =========================================================

    private static async Task<bool> TriggerMacroDroid(
        MimeMessage message,
        string matchedText)
    {
        try
        {
            var payload = new
            {
                event_type = "kabasakal_email",
                sender = message.From.Mailboxes.FirstOrDefault()?.Address,
                subject = message.Subject,
                received = message.Date,
                trigger_text = matchedText
            };

            string json = JsonSerializer.Serialize(payload);

            using var content =
                new StringContent(json, Encoding.UTF8, "application/json");

            Console.WriteLine("[WEBHOOK] Sending...");

            using var response =
                await Http.PostAsync(MacroDroidWebhookUrl, content);

            string responseText = await response.Content.ReadAsStringAsync();

            Console.WriteLine($"[WEBHOOK] HTTP {(int)response.StatusCode}");

            if (!string.IsNullOrWhiteSpace(responseText))
            {
                Console.WriteLine($"[WEBHOOK] {responseText}");
            }

            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[WEBHOOK ERROR] {ex.Message}");
            return false;
        }
    }

}
