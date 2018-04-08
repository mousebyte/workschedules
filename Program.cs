using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Google.Apis.Calendar.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;

namespace workschedule
{
    internal static class Program
    {
        private const string ApplicationName =
            "Dillard's Calendar Import";

        private static readonly string[] Scopes =
        {
            CalendarService.Scope.Calendar,
            GmailService.Scope.GmailModify
        };

        private static readonly Log Log =
            new Log(Console.Out);

        private static readonly Regex Regex = new Regex(
            @"(?<date>\d{2}\/\d{2}): (?<in>\d{2}:\d{2}\w)(?:.*?)(?<out>\d{2}:\d{2}\w)(?= \w{3}\d{3}\s?\r)",
            RegexOptions.Compiled);

        private static async Task Main(string[] args)
        {
            var credPath = Path.Combine(
                Environment.GetFolderPath(Environment
                    .SpecialFolder.Personal),
                $".credentials/{ApplicationName}.json");
            var isTestRun = false;
            var reauthorize = false;
            IEnumerable<Message> messages;
            for (var i = 0; i < args.Length; i++)
                switch (args[i])
                {
                    case "-a":
                        reauthorize = true;
                        break;

                    case "-t":
                        isTestRun = true;
                        Log.WriteAsync(Log.EventType.Info,
                            "TEST RUN");
                        break;

                    case "-p":
                        credPath =
                            Path.GetFullPath(args[++i]);
                        break;
                }
            if (reauthorize)
                Directory.Delete(credPath, true);
            var serviceprovider =
                new ServiceProvider(ApplicationName,
                    credPath, Scopes)
                {
                    Log = Log,
                    MessageQuery =
                        "from:WorkSchedules@dillards.com"
                };

            var parser = new MailParser(Regex)
            {
                DateFormat = "MM/dd hh:mmt",
                ColorId = "2",
                Log = Log
            };


            try
            {
                messages =
                    await serviceprovider.GetMessageAsync(
                        !isTestRun);
            }
            catch (InvalidOperationException)
            {
                Log.WriteAsync(Log.EventType.Info,
                    "Exiting program.");
                return;
            }

            var insertTask = Task.WhenAll(messages
                .SelectMany(parser.ParseMessage)
                .Select(serviceprovider.InsertEventAsync));
            await insertTask;
            Log.WriteAsync(Log.EventType.Info,
                "Finished creating events.");
            if (isTestRun)
                await serviceprovider.EraseEvents();
            Console.Read();
        }
    }
}