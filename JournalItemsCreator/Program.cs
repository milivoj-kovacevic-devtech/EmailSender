using Faker;
using JournalItemsCreator.Shared;
using System;
using System.Collections.Generic;

namespace JournalItemsCreator
{
    class Program
    {
        private static Logger _log;
        private static readonly ConfigManager Config = new ConfigManager();
        private static JournalsController _controller;
        static void Main(string[] args)
        {
            _log = new Logger(Config);
            var mailboxes = Config.GetMailboxes();
            var rnd = new Random();

            _log.Info("Started creating Journal items.");
            while (true)
            {
                foreach (var mailbox in mailboxes)
                {
                    var startTime = rnd.Next(1, 1000);
                    var endTime = startTime + rnd.Next(1, 5);
                    var journalType = GetRandomJournalType(rnd);
                    _controller = new JournalsController(_log, Config)
                    {
                        Subject = TextFaker.Sentence(),
                        Body = TextFaker.Sentences(5),
                        Type = journalType,
                        TypeDescription = journalType,
                        Company = CompanyFaker.Name(),
                        StartTime = startTime,
                        EndTime = endTime
                    };
                    _log.Info(string.Format("Creating journal item [{0}] for mailbox [{1}]", _controller.Type, mailbox));
                    _controller.CreateJournalItem(mailbox);
                }
            }
        }

        private static string GetRandomJournalType(Random rnd)
        {
            var journalTypes = new List<string>() { "Conversation", "Document", "E-mail Message", "Fax", "Letter",
                "Meeting", "Meeting cancellation", "Meeting request", "Meeting response", "Microsoft Excel", "Microsoft Office Access",
                "Microsoft PowerPoint", "Microsoft Word", "Note", "Phone call", "Remote session", "Task", "Task request", "Task response" };

            return journalTypes[rnd.Next(journalTypes.Count)];
        }
    }
}
