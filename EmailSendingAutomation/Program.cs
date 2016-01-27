using System;
using EmailSender.Models;
using EmailSender.Shared;

namespace EmailSender
{

	class Program
	{
        private static readonly ConfigManager Config = new ConfigManager();
		private static SendingController _sendingController;
		private static Logger _log;

		static void Main(string[] args)
		{
			 _log = new Logger(Config);
			_sendingController = new SendingController(Config, _log);
			_log.Info("EmailSendingAutomation init...");

			var rnd = new Random();
			var contactsList = Config.GetContacts();
			var indexLimit = contactsList.Count;

			// Read and delete old messages first
			_sendingController.DeleteOldMailsFromAllMailboxes(contactsList);

			while (true)
			{
				var c = rnd.Next(indexLimit);
				var r = Helper.IncreaseIndex(c, indexLimit);
				var sender = contactsList[c];
				var reply = contactsList[r];
				var reply1 = contactsList[Helper.IncreaseIndex(r, indexLimit)];
				var msgType = (EmailItemType)Enum.GetValues(typeof(EmailItemType)).GetValue(rnd.Next(4));

				switch (msgType)
				{
					// Actions for sending (and replying to) standard textual email messages
					case EmailItemType.Email:
						_sendingController.SendEmail(sender, reply);
						break;

					// Actions for scheduling appointments and responding to meeting requests (accept, decline, accept tentatively)
					case EmailItemType.Meeting:
						_sendingController.ScheduleMeeting(sender, reply, reply1);
						break;

					// Actions for creating a task for current user
					case EmailItemType.Task:
						_sendingController.CreateTask(sender);
						break;
					// Actions for creating journal item for current user
					case EmailItemType.Journal:
						_sendingController.CreateJournal(sender);
						break;
						// If, for some reason, none of the 3 message types are passed to switch statement
					default:
						_log.Error("Unknown message type");
						break;
				}
			}
		}
	}
}