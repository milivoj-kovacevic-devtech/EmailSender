using System.Linq;
using System;
using System.Collections.Generic;
using System.Threading;
using EmailSender.Models;
using Faker;
using Microsoft.Exchange.WebServices.Data;
using Contact = EmailSender.Models.Contact;

namespace EmailSender.Shared
{
	public class SendingController
	{
		private static ConfigManager _config;
		private static Logger _log;

		public SendingController(ConfigManager config, Logger log)
		{
			_config = config;
			_log = log;
		}

		public void SendEmail(Contact senderContact, Contact replyContact)
		{
			var rnd = new Random();
			var attachments = new[] {_config.GetTextAttachmentPath(), _config.GetBinaryAttachmentPath()};
			var currentTime = DateTime.Now;
			var waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
			var timeString = currentTime.ToString();
			var sender = new EmailSender(senderContact.Credentials, _log)
			{
				Subject = "Email " + timeString,
				Body = "This email was sent at: " + timeString,
				ToEmailAddress = replyContact.Username
			};

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

			var reply = new EmailSender(replyContact.Credentials, _log)
			{
				Body = "This is reply to test message sent using EWS Managed API. It was sent at " + currentTime,
				FromEmailAddress = replyContact.Username
			};

			if (rnd.Next(2) == 0)
			{
				reply.AttachmentLocation = _config.GetReplyAttachmentPath();
			}

			reply.Reply(sender.ExtendedProperyDef, sender.TestUniqueId);

			// TODO: Replace these times when shorter periods needed
			var waitAfterReply = DateTime.Now.AddMinutes(1);
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
			while (waitAfterReply < currentTime)
			{
				currentTime = DateTime.Now;
			}

			Console.WriteLine("email message was sent");
		}

		public void ScheduleMeeting(Contact senderContact, Contact replyContact, Contact reply1Contact)
		{
			var meetingReplyValues = Enum.GetValues(typeof (MeetingReplyType));
			var rnd = new Random();
			var recuring = Convert.ToBoolean(rnd.Next(2));
			var currentTime = DateTime.Now;
			var waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
			var timeString = currentTime.ToString();
			var sender = new EmailSender(senderContact.Credentials, _log)
			{
				Subject = "Meeting " + timeString,
				Body = "Test meeting scheduled at: " + timeString,
				StartTime = Helper.GetDateForScheduling(currentTime, rnd),
				Duration = rnd.Next(1, 5)/2,
				Location = "Conf Room First Floor",
				RequiredAttendees = new StringList(),
				OptionalAttendees = new StringList()
			};
			sender.RequiredAttendees.Add(replyContact.Username);
			sender.OptionalAttendees.Add(reply1Contact.Username);

			sender.ScheduleMeeting(recuring);

			while (waitAfterSending > currentTime)
			{
				currentTime = DateTime.Now;
				Thread.Sleep(10000);
			}

			timeString = currentTime.ToString();
			var reply = new EmailSender(replyContact.Credentials, _log)
			{
				Body = "Reply from required attendee to test meeting sent at: " + timeString
			};
			var meetingReplyType = (MeetingReplyType) meetingReplyValues.GetValue(rnd.Next(3));
			reply.ReplyToMeetingRequest(sender.Subject, meetingReplyType);

			var reply1 = new EmailSender(reply1Contact.Credentials, _log)
			{
				Body = "Reply from optional attendee to test meeting sent at " + timeString
			};
			meetingReplyType = (MeetingReplyType) meetingReplyValues.GetValue(rnd.Next(3));
			reply1.ReplyToMeetingRequest(sender.Subject, meetingReplyType);

			// TODO: Replace these times when shorter periods needed
			var waitAfterReply = DateTime.Now.AddMinutes(1);
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
			while (waitAfterReply < currentTime)
			{
				currentTime = DateTime.Now;
			}

			Console.WriteLine("Meeting has been scheduled");
		}

		public void CreateTask(Contact senderContact)
		{
			var rnd = new Random();
			var recuring = Convert.ToBoolean(rnd.Next(2));
			var currentTime = DateTime.Now;
			var taskSubject = "Task subject " + currentTime;
			var taskBody = "This task was created at " + currentTime;
			var sender = new EmailSender(senderContact.Credentials, _log)
			{
				StartTime = Helper.GetDateForScheduling(currentTime, rnd),
				Duration = rnd.Next(1, 6)
			};
			const int reminder = 15;
			sender.CreateTask(taskSubject, taskBody, reminder, recuring);

			Console.WriteLine("Task with subject {0} has been created", taskSubject);
		}

		public void DeleteOldMailsFromAllMailboxes(List<Contact> contactsList)
		{
			foreach (var messageReader in contactsList.Select(contact => new EmailSender(contact.Credentials, _log)))
			{
				messageReader.ReadOldMessages(_config.GetDeleteOldMailFlag());
			}
		}

		public void CreateJournal(Contact journalSender)
		{
			_log.Info(journalSender.Username);
			var rnd = new Random();
			var startTime = rnd.Next(1, 1000);
			var endTime = startTime + rnd.Next(1, 5);
			var journalType = GetRandomJournalType(rnd);
			var subject = TextFaker.Sentence();
			var body = TextFaker.Sentences(5);
			var companies = GetCompanies(rnd);
			var journal = new Journal()
			{
				Subject = subject,
				Body = body,
				Type = journalType.Type,
				TypeDescription = journalType.Type,
				IconIndex = journalType.IconIndex,
				Companies = companies,
				StartTime = DateTime.Now.AddHours(startTime),
				EndTime = DateTime.Now.AddHours(endTime)
			};
			var sender = new EmailSender(journalSender.Credentials, _log);
			sender.CreateJournalItem(journal);

			_log.Info(string.Format("Journal with subject [{0}] created for mailbox [{1}]", subject, journalSender.Username));

		}

		private static JournalType GetRandomJournalType(Random rnd)
		{
			var journalTypes = new List<JournalType>();
			journalTypes.Add(new JournalType("Conversation", 1537));
			journalTypes.Add(new JournalType("Document", 1554));
			journalTypes.Add(new JournalType("E-mail Message", 1538));
			journalTypes.Add(new JournalType("Fax", 1545));
			journalTypes.Add(new JournalType("Letter", 1548));
			journalTypes.Add(new JournalType("Meeting", 1555));
			journalTypes.Add(new JournalType("Meeting cancellation", 1556));
			journalTypes.Add(new JournalType("Meeting request", 1539));
			journalTypes.Add(new JournalType("Meeting response", 1540));
			journalTypes.Add(new JournalType("Microsoft Excel", 1550));
			journalTypes.Add(new JournalType("Microsoft Office Access", 1552));
			journalTypes.Add(new JournalType("Microsoft PowerPoint", 1551));
			journalTypes.Add(new JournalType("Microsoft Word", 1549));
			journalTypes.Add(new JournalType("Note", 1544));
			journalTypes.Add(new JournalType("Phone call", 1546));
			journalTypes.Add(new JournalType("Remote session", 1557));
			journalTypes.Add(new JournalType("Task", 1547));
			journalTypes.Add(new JournalType("Task request", 1542));
			journalTypes.Add(new JournalType("Task response", 1543));

			return journalTypes[rnd.Next(journalTypes.Count)];
		}

		private static string[] GetCompanies(Random rnd)
		{
			var numOfCompanies = rnd.Next(1, 3);
			var retVal = new string[numOfCompanies];
			for (var i = 0; i < numOfCompanies; i++)
			{
				retVal[i] = CompanyFaker.Name();
			}
			return retVal;
		}
	}
}
