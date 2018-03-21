using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Util.Store;

namespace workschedule
{
    internal static class Program
    {
        private const string ApplicationName = "Dillard's Calendar Import";
        private static readonly string[] Scopes = {CalendarService.Scope.Calendar, GmailService.Scope.GmailModify};
        private static readonly Log Log = new Log(Console.Out);

        private static readonly Regex Regex = new Regex(
            @"(?<date>\d{2}\/\d{2}): (?<in>\d{2}:\d{2}\w)(?:.*?)(?<out>\d{2}:\d{2}\w)(?= \w{3}\d{3}\s?\r)",
            RegexOptions.Compiled);


        private static async Task Main(string[] args)
        {
            var credPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                $".credentials/{ApplicationName}.json");
            var isTestRun = false;
            var reauthorize = false;
            for (var i = 0; i < args.Length; i++)
                switch (args[i])
                {
                    case "-a":
                        reauthorize = true;
                        break;
                    case "-t":
                        isTestRun = true;
                        Log.WriteAsync(Log.EventType.Info, "TEST RUN");
                        break;
                    case "-p":
                        credPath = Path.GetFullPath(args[++i]);
                        break;
                }
            var parser = new MailParser(ApplicationName, Regex, await GetCredentials(credPath, reauthorize))
            {
                DateFormat = "MM/dd hh:mmt",
                MessageQuery = "from:WorkSchedules@dillards.com",
                ColorId = "2",
                Log = Log
            };
            await parser.PopulateCalendarAsync();
            if (isTestRun) await parser.EraseEvents();
        }

        private static async Task<UserCredential> GetCredentials(string path, bool reauthorize = false)
        {
            if (reauthorize && Directory.Exists(path)) Directory.Delete(path, true);
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                    Scopes, "user", CancellationToken.None, new FileDataStore(path, true));
            }

            return credential;
        }
    }
}