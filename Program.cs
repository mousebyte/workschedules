using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace workschedule
{
    internal static class Program
    {
        private const string ApplicationName = "Dillard's Calendar Import";
        private static readonly string[] Scopes = {CalendarService.Scope.Calendar, GmailService.Scope.GmailModify};

        private static readonly Regex Regex = new Regex(
            @"(\w{3})\s*(?<date>\d{2}\/\d{2}): (?<in>\d{2}:\d{2}\w)(?:.*?)(?<out>\d{2}:\d{2}\w) \w{3}\d{3}\s?\r",
            RegexOptions.Compiled);

        private static void Main()
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

            var calendarService = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            var mailService = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            var msgRequest = new UsersResource.MessagesResource.ListRequest(mailService, "me")
            {
                Q = "from:WorkSchedules@dillards.com",
                LabelIds = "INBOX"
            };

            //Execute the message request, make sure there's exactly one viable message to work with.
            var msgResponse = msgRequest.Execute().Messages;
            if (msgResponse == null || msgResponse.Count > 1) return;
            var msg = msgResponse[0];

            var body = ConvertPayloadBody(mailService.Users.Messages.Get("me", msg.Id).Execute().Payload.Body.Data);
            if (body.StartsWith("JANUARY")) return; //No code to handle this right now, so exit the program.
            var year = body.Substring(body.IndexOf(' ') + 1, 4);
            var matches = Regex.Matches(body);
            Console.WriteLine("Populating calendar...");

            foreach (Match matchItem in matches)
            {
                var calEvent = calendarService.Events.Insert(new Event
                {
                    Summary = "Work",
                    ColorId = "2",
                    Start = new EventDateTime
                    {
                        DateTime = DateTime.Parse($"{matchItem.Groups["date"]}/{year} {matchItem.Groups["in"]}M")
                    },
                    End = new EventDateTime
                    {
                        DateTime = DateTime.Parse($"{matchItem.Groups["date"]}/{year} {matchItem.Groups["out"]}M")
                    }
                }, "primary").Execute();
                Console.WriteLine($"Created work entry on {calEvent.Start.DateTime?.ToShortDateString()}");
            }

            //Remove the inbox label from the message when done so that it isn't seen next time the program runs.
            //This has the same effect as archiving the message in gmail.
            mailService.Users.Messages.Modify(new ModifyMessageRequest
            {
                RemoveLabelIds = new[] {"INBOX"}
            }, "me", msg.Id).Execute();
            Console.WriteLine($"Done! Created {matches.Count} events.");
            Console.Read();
        }

        private static string ConvertPayloadBody(string data)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(data.Replace('-', '+')));
        }
    }
}