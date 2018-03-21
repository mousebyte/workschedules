using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Http;
using Google.Apis.Services;

namespace workschedule
{
    public class MailParser
    {
        private readonly CalendarService _calendarService;
        private readonly GmailService _mailService;
        private readonly Regex _regex;
        private readonly List<Event> _insertedEvents = new List<Event>();

        public MailParser(string applicationName, Regex regex, IConfigurableHttpClientInitializer credential)
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

        public string UserId { get; set; }

        public string CalendarId { get; set; }

        public string MessageQuery { get; set; }

        public string DateFormat { get; set; }

        public string Summary { get; set; }

        public string ColorId { get; set; }

        public Log Log { private get; set; }

        public async Task<bool> PopulateCalendarAsync()
        {
            Log.WriteAsync(Log.EventType.Info, "Requesting messages...", $"User ID: {UserId}",
                $"Query: {MessageQuery}");
            var msgResponse = await new UsersResource.MessagesResource.ListRequest(_mailService, UserId)
            {
                Q = MessageQuery,
                LabelIds = "INBOX"
            }.ExecuteAsync();
            Log.WriteAsync(Log.EventType.Info, "Response recieved.");

            Message msg;
            try
            {
                msg = msgResponse.Messages.Single();
                Log.WriteAsync(Log.EventType.Info, $"Source message {msg.Id} found.");
            }
            catch (InvalidOperationException)
            {
                Log.WriteAsync(Log.EventType.Error,
                    "Failed to retrieve source message: Query returned more than one match.",
                    $"Match Count: {msgResponse.Messages.Count}",
                    $"Match IDs: {msgResponse.Messages.Aggregate("", (s, message) => $"{s}, {message.Id}")}");
                return false;
            }
            catch (ArgumentNullException)
            {
                Log.WriteAsync(Log.EventType.Error, "Failed to retrieve source message: Query returned no matches.");
                return false;
            }

            //Get the body of the message as a string and find all matches
            Log.WriteAsync(Log.EventType.Info, "Converting message body...");
            var body = ConvertPayloadBody(_mailService.Users.Messages.Get(UserId, msg.Id).Execute());
            Log.WriteAsync(Log.EventType.Info, "Message body converted.");

            _insertedEvents.Clear();
            Log.WriteAsync(Log.EventType.Info, "Starting event insertion...");
            var insertTask =
                Task.WhenAll(
                    _regex.Matches(body).Cast<Match>().Select(async m => await InsertEventAsync(ParseMatch(m))));


            _mailService.Users.Messages.Modify(new ModifyMessageRequest
            {
                RemoveLabelIds = new[] {"INBOX"}
            }, UserId, msg.Id).Execute();
            Log.WriteAsync(Log.EventType.Info, "Source message sent to archive.");

            await insertTask;
            Log.WriteAsync(Log.EventType.Info, "Finished creating events.");
            return true;
        }

        public async Task EraseEvents()
        {
            if (_insertedEvents.Count == 0) return;
            Log.WriteAsync(Log.EventType.Info, "Erasing events from previous run...");
            await Task.WhenAll(_insertedEvents.Select(async e =>
            {
                await _calendarService.Events.Delete(CalendarId, e.Id).ExecuteAsync();
                Log.WriteAsync(Log.EventType.Info, $"Erased event {e.Id}.");
            }));
            Log.WriteAsync(Log.EventType.Info, "Finished erasing events.");
        }

        private async Task InsertEventAsync(Event e)
        {
            var response = await _calendarService.Events.Insert(e, CalendarId).ExecuteAsync();
            _insertedEvents.Add(response);
            Log.WriteAsync(Log.EventType.Info, $"Created event '{response.Summary}'",
                $"Date: {response.Start.DateTime:d}",
                $"Time: {response.Start.DateTime:hh:mm tt}-{response.End.DateTime:hh:mm tt}");
        }

        private Event ParseMatch(Match match)
        {
            if (!DateTime.TryParseExact($"{match.Groups["date"]} {match.Groups["in"]}", DateFormat, null,
                    DateTimeStyles.None, out var start) || !DateTime.TryParseExact(
                    $"{match.Groups["date"]} {match.Groups["out"]}", DateFormat, null,
                    DateTimeStyles.None, out var end))
            {
                Log.WriteAsync(Log.EventType.Error, "Error parsing DateTime.", $"Match: {match.Value}",
                    $"Format: {DateFormat}");
                throw new InvalidOperationException("Error parsing DateTime.");
            }

            Log.WriteAsync(Log.EventType.Info, "Date parsed successfully.");

            if (DateTime.Now.Month == 12 && start.Month == 1)
            {
                Log.WriteAsync(Log.EventType.Info, "Adjusting year of parsed date.");
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
            return Encoding.UTF8.GetString(Convert.FromBase64String(msg.Payload.Body.Data.Replace('-', '+')));
        }
    }
}