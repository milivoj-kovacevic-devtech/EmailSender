using Independentsoft.Exchange;
using JournalItemsCreator.Shared;
using System;
using System.Net;

namespace JournalItemsCreator
{
    public class JournalsController
    {
        private static Logger _log;
        private static ConfigManager _config;
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Type { get; set; }
        public string TypeDescription { get; set; }
        public string Company { get; set; }
        public int StartTime { get; set; }
        public int EndTime { get; set; }

        public JournalsController(Logger log, ConfigManager config)
        {
            _log = log;
            _config = config;
        }

        public void CreateJournalItem(string mailboxUsername)
        {
            var credential = new NetworkCredential(mailboxUsername, _config.GetMailboxPassword());
            var service = new Service(_config.GetEwsUrl(), credential);
            try
            {
	            var journal = new Journal
	            {
		            Subject = Subject,
		            Body = new Body(Body),
		            Type = Type,
		            TypeDescription = TypeDescription
	            };
	            journal.Companies.Add(Company);
                journal.StartTime = DateTime.Today.AddHours(StartTime);
                journal.EndTime = DateTime.Today.AddHours(EndTime);
				
                var itemId = service.CreateItem(journal, StandardFolder.Journal);

                var testis = journal.ExtendedProperties;
                Console.WriteLine(testis);
            }
            catch (ServiceRequestException ex)
            {
                _log.Error(string.Format("Error: [{0}]", ex.Message));
                _log.Error(string.Format("Error xml output: \n{0}", ex.XmlMessage));
                Console.Read();
            }
            catch (WebException ex)
            {
                _log.Error(string.Format("Error: [{0}]", ex.Message));
                Console.Read();
            }
        }
    }
}
