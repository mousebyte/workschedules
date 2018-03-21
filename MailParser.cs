using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;

namespace workschedule
{
    class Log
    {
        public enum EventType
        {
            Error,
            Info
        }
        private TextWriter _writer;

        public Log(TextWriter writer)
        {
            _writer = writer;
            _writer.WriteLine($"Begin log {DateTime.Now:G}");
        }

        public async void WriteAsync(string s, EventType eventType)
        {
            await _writer.WriteLineAsync($"    [{DateTime.Now:T}] {eventType:G} - {s}");
        }
    }
    class MailParser
    {
        private readonly Regex _regex;

        private readonly CalendarService _calendarService;

        private readonly GmailService _mailService;

        public string UserId { get; set; }

        public string CalendarId { get; set; }

        public string MessageQuery { get; set; }

        public string DateFormat { get; set; }

        public string Summary { get; set; }

        public string ColorId { get; set; }

        public Log Log { private get; set; }

        public MailParser(string applicationName, Regex regex, UserCredential credential)
        {
            _regex = regex;
            _calendarService = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });
            _mailService = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            });
            UserId = "me";
            CalendarId = "primary";
            Summary = "Work";
            ColorId = "1";
        }

        public async Task<bool> PopulateCalendarAsync()
        {
            var msgResponse = new UsersResource.MessagesResource.ListRequest(_mailService, UserId)
            {
                Q = MessageQuery, //"from:WorkSchedules@dillards.com",
                LabelIds = "INBOX"
            }.Execute();

            Message msg;
            try
            {
                msg = msgResponse.Messages.Single();
            }
            catch (InvalidOperationException)
            {
                Log.WriteAsync("Failed to retrieve source message: Query did not return exactly 1 match.", Log.EventType.Error);
                return false;
            }

            //Get the body of the message as a string and find all matches
            var body = ConvertPayloadBody(_mailService.Users.Messages.Get(UserId, msg.Id).Execute());

            var insertTask =
                Task.WhenAll(
                    _regex.Matches(body).Cast<Match>().Select(async m => await InsertEventAsync(ParseMatch(m))));

            await _mailService.Users.Messages.Modify(new ModifyMessageRequest
            {
                RemoveLabelIds = new[] {"INBOX"}
            }, UserId, msg.Id).ExecuteAsync();

            await insertTask;
            Log.WriteAsync("Finished creating events.", Log.EventType.Info);
            return true;
        }

        private async Task InsertEventAsync(Event e)
        {
            await _calendarService.Events.Insert(e, CalendarId).ExecuteAsync();
            Log.WriteAsync($"Created event '{e.Summary} {e.Start.DateTime:d}'", Log.EventType.Info);
        }

        private Event ParseMatch(Match match)
        {
            DateTime.TryParseExact($"{match.Groups["date"]} {match.Groups["in"]}", DateFormat, null,
                DateTimeStyles.None, out var start);
            DateTime.TryParseExact($"{match.Groups["date"]} {match.Groups["out"]}", DateFormat, null,
                DateTimeStyles.None, out var end);

            if (DateTime.Now.Month == 12 && start.Month == 1)
            {
                start = start.AddYears(1);
                end = end.AddYears(1);
            }

            return new Event
            {
                Summary = Summary,
                ColorId = ColorId,
                Start = new EventDateTime
                {
                    DateTime = start
                },
                End = new EventDateTime
                {
                    DateTime = end
                }
            };
        }

        private static string ConvertPayloadBody(Message msg)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(msg.Payload.Body
                .Data.Replace('-', '+')));
        }
    }
}
