﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace workschedule
{
    internal class ServiceProvider : IDisposable
    {
        private readonly ModifyMessageRequest
            _archiveRequest = new ModifyMessageRequest
            {
                RemoveLabelIds = new[] {"INBOX"}
            };

        private readonly CalendarService _calendarService;

        private readonly List<Event> _insertedEvents =
            new List<Event>();

        private readonly GmailService _mailService;

        private string _calendarId;

        private EventsResource.ListRequest
            _checkEventsRequest;

        private string _messageQuery;

        private UsersResource.MessagesResource.ListRequest
            _messageRequest;

        private string _userId;

        public ServiceProvider(string applicationName,
            string path, params string[] scopes)
        {
            _userId = "me";
            _calendarId = "primary";
            UserCredential credential;
            using (var stream = new FileStream(
                "client_secret.json", FileMode.Open,
                FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker
                    .AuthorizeAsync(
                        GoogleClientSecrets.Load(stream)
                            .Secrets,
                        scopes, "user",
                        CancellationToken.None,
                        new FileDataStore(path, true))
                    .Result;
            }

            _calendarService = new CalendarService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = applicationName
                });
            _mailService = new GmailService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = applicationName
                });
            _checkEventsRequest =
                new EventsResource.ListRequest(
                    _calendarService, CalendarId)
                {
                    PrivateExtendedProperty =
                        "AutoGenerated=true",
                    MaxResults = 1
                };
            _messageRequest =
                new UsersResource.MessagesResource.
                    ListRequest(_mailService, UserId)
                    {
                        Q = MessageQuery,
                        LabelIds = "INBOX"
                    };
        }

        public string CalendarId
        {
            get => _calendarId;
            set
            {
                _calendarId = value;
                _checkEventsRequest =
                    new EventsResource.ListRequest(
                        _calendarService, value);
            }
        }

        public Log Log { private get; set; }

        public string MessageQuery
        {
            get => _messageQuery;
            set
            {
                _messageQuery = value;
                _messageRequest.Q = value;
            }
        }

        public string UserId
        {
            get => _userId;
            set
            {
                _userId = value;
                _messageRequest =
                    new UsersResource.MessagesResource.
                        ListRequest(_mailService, value);
            }
        }

        public async Task EraseEvents()
        {
            if (_insertedEvents.Count == 0) return;
            Log.WriteAsync(Log.EventType.Info,
                "Erasing events from previous run...");
            await Task.WhenAll(_insertedEvents.Select(
                async e =>
                {
                    await _calendarService.Events
                        .Delete(CalendarId, e.Id)
                        .ExecuteAsync();
                    Log.WriteAsync(Log.EventType.Info,
                        $"Erased event {e.Id}.");
                }));
            Log.WriteAsync(Log.EventType.Info,
                "Finished erasing events.");
            _insertedEvents.Clear();
        }

        public async Task<IEnumerable<Message>>
            GetMessageAsync(bool archiveMessages = true)
        {
            Log.WriteAsync(Log.EventType.Info,
                "Requesting messages...",
                $"User ID: {UserId}",
                $"Query: {MessageQuery}");
            var requestTask =
                _messageRequest.ExecuteAsync();
            var msgResponse = await requestTask;
            Log.WriteAsync(Log.EventType.Info,
                "Response recieved.");

            if (msgResponse.Messages.Count == 0)
            {
                Log.WriteAsync(Log.EventType.Error,
                    "No matching messages found.");
                throw new InvalidOperationException(
                    "Query returned no matching messages.");
            }

            if (archiveMessages)
                foreach (var msg in msgResponse.Messages)
                    _mailService.Users.Messages
                        .Modify(_archiveRequest, UserId,
                            msg.Id).Execute();

            return await Task.WhenAll(
                msgResponse.Messages.Select(async m =>
                    await _mailService
                        .Users.Messages
                        .Get(UserId, m.Id).ExecuteAsync()));
        }

        public async Task InsertEventAsync(Event e)
        {
            _checkEventsRequest.TimeMax =
                e.Start.DateTime?.AddMinutes(1);
            if ((await _checkEventsRequest.ExecuteAsync())
                .Items.Count == 1)
            {
                Log.WriteAsync(Log.EventType.Info,
                    $"Work already scheduled on {e.Start.DateTime?.Date:d}.");
                return;
            }

            var response = _calendarService.Events
                .Insert(e, CalendarId).ExecuteAsync();
            Log.WriteAsync(Log.EventType.Info,
                $"Creating event '{e.Summary}'",
                $"Date: {e.Start.DateTime:d}",
                $"Time: {e.Start.DateTime:hh:mm tt}-{e.End.DateTime:hh:mm tt}");
            _insertedEvents.Add(await response);
        }

        public void Dispose()
        {
            _calendarService.Dispose();
            _mailService.Dispose();
        }
    }
}