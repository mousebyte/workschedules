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

        private static readonly Regex Regex = new Regex(
            @"(?<date>\d{2}\/\d{2}): (?<in>\d{2}:\d{2}\w)(?:.*?)(?<out>\d{2}:\d{2}\w)(?= \w{3}\d{3}\s?\r)",
            RegexOptions.Compiled);

        private static async Task Main()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                var credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/dillards-calendar-import.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                    Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                Console.WriteLine($"Credential file saved to: {credPath}");
            }
            var parser = new MailParser(ApplicationName, Regex, credential)
            {
                DateFormat = "MM/dd hh:mmt",
                MessageQuery = "from:WorkSchedules@dillards.com",
                ColorId = "2",
                Log = new Log(Console.Out)
            };
            await parser.PopulateCalendarAsync();
            Console.Read();
        }
    }
}