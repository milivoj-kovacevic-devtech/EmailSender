using System;
using log4net;
using log4net.Config;
using EmailSender.Shared;

namespace EmailSender
{

	class Program
	{
		private static readonly SendingController SendingController = new SendingController();
		private static readonly ILog Log = LogManager.GetLogger(typeof(Program));

		static void Main(string[] args)
		{

			// TODO Check this error
			// Unhandled Exception: System.NullReferenceException: Object reference not set to an instance of an object.
			// at EmailSender.Shared.SendingController.ScheduleMeeting(Contact senderContact
			// , Contact replyContact, Contact reply1Contact) in d:\Projects\ParallelsMigrationTool\EmailSender\EmailSendingAutomation\Shared\SendingController.cs:line 78
			// at EmailSender.Program.Main(String[] args) in d:\Projects\ParallelsMigrationTool\EmailSender\EmailSendingAutomation\Program.cs:line 46

			// log4net configurator
			XmlConfigurator.Configure();
			Log.Info("EmailSendingAutomation init...");

			var rnd = new Random();
			var contactsList = ConfigManager.GetContacts();
			var indexLimit = contactsList.Count;

			// Read and delete old messages first
			SendingController.DeleteOldMailsFromAllMailboxes(contactsList);

			while (true)
			{
				var c = rnd.Next(indexLimit);
				var r = Helper.IncreaseIndex(c, indexLimit);
				var sender = contactsList[c];
				var reply = contactsList[r];
				var reply1 = contactsList[Helper.IncreaseIndex(r, indexLimit)];
				var msgType = (EmailItemType)Enum.GetValues(typeof(EmailItemType)).GetValue(rnd.Next(3));

				switch (msgType)
				{
					// Actions for sending (and replying to) standard textual email messages
					case EmailItemType.Email:
						SendingController.SendEmail(sender, reply);
						break;

					// Actions for scheduling appointments and responding to meeting requests (accept, decline, accept tentatively)
					case EmailItemType.Meeting:
						SendingController.ScheduleMeeting(sender, reply, reply1);
						break;

					// Actions for creating a task for current user
					case EmailItemType.Task:
						SendingController.CreateTask(sender);
						break;

					// If, for some reason, none of the 3 message types are passed to switch statement
					default:
						Log.Error("Unknown message type");
						break;
				}
			}
		}
	}
}
