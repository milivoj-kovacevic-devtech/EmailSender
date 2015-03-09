using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using log4net;
using log4net.Config;
using System.Collections.Specialized;

namespace EmailSender
{

	class Program
	{
		private static readonly ILog Log = LogManager.GetLogger(typeof(Program));

		// Increases array index, when limit is reached next value is 0 (zero)
		public static int IncreaseIndex(int index, int limit)
		{
			int incIndex;
			if (index == limit - 1)
			{
				incIndex = 0;
			}
			else
			{
				incIndex = index + 1;
			}

			return incIndex;
		}

		// Gets date for scheduling meetings or creating tasks (1-3 days from today, at 9-16 o'clock)
		public static DateTime GetDateForScheduling(DateTime currentTime, Random rnd)
		{
			DateTime date = currentTime.AddDays(rnd.Next(1, 3));

			return new DateTime(date.Year, date.Month, date.Day, rnd.Next(9, 17), 0, 0);
		}

		static void Main(string[] args)
		{

			// log4net configurator
			XmlConfigurator.Configure();
			Log.Info("EmailSendingAutomation init...");

			// The ammount of time for the application to run
			var timeToRun = ConfigManager.HoursToWork;

			var currentTime = DateTime.Now;
			var timeToStop = DateTime.Now.AddHours(timeToRun);

			var rnd = new Random();

			// Used for storing current time as string for email subject and body
			string timeString;

			// Sets recurence for meetings and tasks
			var recuring = false;

			EmailItemType msgType = EmailItemType.Email;
			MeetingReplyType meetingReplyType = MeetingReplyType.Accept;

			Array itemTypeValues = Enum.GetValues(typeof(EmailItemType));
			Array meetingReplyValues = Enum.GetValues(typeof(MeetingReplyType));

			var contactsList = ConfigManager.GetContacts();

			int indexLimit = contactsList.Count;

			String[] attachments = new String[2];
			attachments[0] = ConfigManager.TextAttachment;
			attachments[1] = ConfigManager.BinaryAttachment;

			// Read and delete old messages first
			foreach (var contact in contactsList)
			{
				EmailSender messageReader = new EmailSender(contact.Credentials);
				messageReader.ReadOldMessages(ConfigManager.DeleteWeekOld);
			}

			while (currentTime < timeToStop)
			{
				// TODO: Replace these times when shorter periods needed
				//DateTime waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
				DateTime waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait to reply to email
				DateTime waitAfterReply;
				int c = rnd.Next(indexLimit);
				int r = IncreaseIndex(c, indexLimit);
				int r1 = IncreaseIndex(r, indexLimit);

				WebCredentials senderCredentials = contactsList[c].Credentials;
				WebCredentials recieverCredentials = contactsList[r].Credentials;
				WebCredentials reciever1Credentials = contactsList[r1].Credentials;

				EmailSender sender = new EmailSender(senderCredentials);
				EmailSender reply = new EmailSender(recieverCredentials);
				EmailSender reply1 = new EmailSender(reciever1Credentials);

				msgType = (EmailItemType)itemTypeValues.GetValue(rnd.Next(3));

				switch (msgType)
				{
					// Actions for sending (and replying to) standard textual email messages
					case EmailItemType.Email:

						timeString = currentTime.ToString();
						sender.Subject = "Email " + timeString;
						sender.Body = "This email was sent at: " + timeString;
						sender.ToEmailAddress = contactsList[r].EmailAddress;

						if (rnd.Next(2) == 0)
						{
							sender.AttachmentLocation = attachments[rnd.Next(2)];
						}

						sender.SendMessage();

						while (waitAfterSending > currentTime)
						{
							currentTime = DateTime.Now;
							Thread.Sleep(10000);
						}


						reply.Body = "This is reply to test message sent using EWS Managed API. It was sent at " + currentTime;
						reply.FromEmailAddress = contactsList[c].EmailAddress;

						if (rnd.Next(2) == 0)
						{
							reply.AttachmentLocation = ConfigManager.ReplyAttachment;
						}

						reply.Reply(sender.ExtendedProperyDef, sender.TestUniqueId);

						// TODO: Replace these times when shorter periods needed
						//waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
						waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
						while (waitAfterReply < currentTime)
						{
							currentTime = DateTime.Now;
						}


						Console.WriteLine("email message was sent");
						break;

					// Actions for scheduling appointments and responding to meeting requests (accept, decline, accept tentatively)
					case EmailItemType.Meeting:
						timeString = currentTime.ToString();
						sender.Subject = "Meeting " + timeString;
						sender.Body = "Test meeting scheduled at: " + timeString;
						sender.StartTime = GetDateForScheduling(currentTime, rnd);
						sender.Duration = rnd.Next(1, 5) / 2;
						sender.Location = "Conf Room First Floor";
						StringList requiredAttendees = new StringList();
						requiredAttendees.Add(contactsList[r].EmailAddress);
						sender.RequiredAttendees = requiredAttendees;

						StringList optionalAttendees = new StringList();
						optionalAttendees.Add(contactsList[r1].EmailAddress);
						sender.OptionalAttendees = optionalAttendees;

						recuring = Convert.ToBoolean(rnd.Next(2));
						sender.ScheduleMeeting(recuring);

						while (waitAfterSending > currentTime)
						{
							currentTime = DateTime.Now;
							Thread.Sleep(10000);
						}

						timeString = currentTime.ToString();
						reply.Body = "Reply from required attendee to test meeting sent at: " + timeString;
						meetingReplyType = (MeetingReplyType)meetingReplyValues.GetValue(rnd.Next(3));
						reply.ReplyToMeetingRequest(sender.Subject, meetingReplyType);

						reply1.Body = "Reply from optional attendee to test meeting sent at " + timeString;
						meetingReplyType = (MeetingReplyType)meetingReplyValues.GetValue(rnd.Next(3));
						reply1.ReplyToMeetingRequest(sender.Subject, meetingReplyType);

						// TODO: Replace these times when shorter periods needed
						//waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
						waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
						while (waitAfterReply < currentTime)
						{
							currentTime = DateTime.Now;
						}


						Console.WriteLine("Meeting has been scheduled");
						break;

					// Actions for creating a task for current user
					case EmailItemType.Task:
						string taskSubject = "Task subject " + currentTime;
						string taskBody = "This task was created at " + currentTime;
						sender.StartTime = GetDateForScheduling(currentTime, rnd);
						sender.Duration = rnd.Next(1, 6);
						int reminder = 15;
						recuring = Convert.ToBoolean(rnd.Next(2));
						sender.CreateTask(taskSubject, taskBody, reminder, recuring);

						Console.WriteLine("Task with subject " + taskSubject + " has been created");
						break;

					// If, for some reason, none of the 3 message types are passed to switch statement
					default:
						Console.WriteLine("Unknown message type");
						break;
				}

				currentTime = DateTime.Now;
			}

			Log.Info("EmailSendingAutomation end.");
		}
	}
}
