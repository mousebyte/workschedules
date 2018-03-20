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
    class MailParser
    {
        private readonly Regex _regex;

        private readonly UserCredential _credential;

        private readonly CalendarService _calendarService;

        private readonly GmailService _mailService;

        public string UserId { get; set; }

        public string CalendarId { get; set; }

        public string MessageQuery { get; set; }

        public string DateFormat { get; set; }

        public string Summary { get; set; }

        public string ColorId { get; set; }

        public TextWriter Log { private get; set; }

        public MailParser(string applicationName, Regex regex, UserCredential credential)
        {
            _regex = regex;
            _credential = credential;
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

        public bool PopulateCalendar()
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
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e);
                throw;
            }

            //Get the body of the message as a string and find all matches
            var body = MailParser.ConvertPayloadBody(_mailService.Users.Messages.Get(UserId, msg.Id).Execute());
            var matches = _regex.Matches(body);

            //Insert a new calendar entry for each match
            foreach (Match match in matches)
                _calendarService.Events.Insert(ParseMatch(match), CalendarId).Execute();

            //Remove the inbox label from the message when done so that it isn't seen next time the program runs.
            //This has the same effect as archiving the message in gmail.
            _mailService.Users.Messages.Modify(new ModifyMessageRequest
            {
                RemoveLabelIds = {"INBOX"}
            }, UserId, msg.Id);
            return true;
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
