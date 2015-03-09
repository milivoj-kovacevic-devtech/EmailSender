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
        private const string UnknownMessage = "Unknown message type";

		static void Main(string[] args)
		{

			// log4net configurator
			XmlConfigurator.Configure();
			Log.Info("EmailSendingAutomation init...");

            var currentTime = DateTime.Now;
			var timeToStop = DateTime.Now.AddHours(ConfigManager.HoursToWork);

			var rnd = new Random();

		    var itemTypeValues = Enum.GetValues(typeof(EmailItemType));

			var contactsList = ConfigManager.GetContacts();

			var indexLimit = contactsList.Count;

			var attachments = new string[2] {ConfigManager.TextAttachment, ConfigManager.BinaryAttachment};

			// Read and delete old messages first
			SendingController.DeleteOldMailsFromAllMailboxes(contactsList);

		    while (currentTime < timeToStop)
			{
				// TODO: Replace these times when shorter periods needed
				//DateTime waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
				var waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait to reply to email
				var c = rnd.Next(indexLimit);
				var r = Helper.IncreaseIndex(c, indexLimit);
				var r1 = Helper.IncreaseIndex(r, indexLimit);

				var sender = new EmailSender(contactsList[c].Credentials);
				var reply = new EmailSender(contactsList[r].Credentials);
				var reply1 = new EmailSender(contactsList[r1].Credentials);

				var msgType = (EmailItemType)itemTypeValues.GetValue(rnd.Next(3));

			    switch (msgType)
				{
					// Actions for sending (and replying to) standard textual email messages
					case EmailItemType.Email:
						SendingController.SendEmail(sender, reply, contactsList[r], waitAfterSending, currentTime, attachments);
						break;

					// Actions for scheduling appointments and responding to meeting requests (accept, decline, accept tentatively)
					case EmailItemType.Meeting:
						SendingController.ScheduleMeeting(sender, reply, reply1, contactsList[r], waitAfterSending, currentTime);
						break;

					// Actions for creating a task for current user
					case EmailItemType.Task:
						SendingController.CreateTask(sender, currentTime);
						break;

					// If, for some reason, none of the 3 message types are passed to switch statement
					default:
				        Log.Error(UnknownMessage);
				        Console.WriteLine(UnknownMessage);
						break;
				}

				currentTime = DateTime.Now;
			}

			Log.Info("EmailSendingAutomation end.");
		}
	}
}
