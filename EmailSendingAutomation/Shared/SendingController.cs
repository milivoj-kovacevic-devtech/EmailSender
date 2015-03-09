using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading;

namespace EmailSender.Shared
{
	public class SendingController
	{
		public void SendEmail(EmailSender sender, EmailSender reply, Contact contact, DateTime waitAfterSending, DateTime currentTime, string[] attachments)
		{
			var rnd = new Random();
			var timeString = currentTime.ToString();
			sender.Subject = "Email " + timeString;
			sender.Body = "This email was sent at: " + timeString;
			sender.ToEmailAddress = contact.EmailAddress;

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
			reply.FromEmailAddress = contact.EmailAddress;

			if (rnd.Next(2) == 0)
			{
				reply.AttachmentLocation = ConfigManager.ReplyAttachment;
			}

			reply.Reply(sender.ExtendedProperyDef, sender.TestUniqueId);

			// TODO: Replace these times when shorter periods needed
			var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
			while (waitAfterReply < currentTime)
			{
				currentTime = DateTime.Now;
			}

			Console.WriteLine("email message was sent");
		}

		public void ScheduleMeeting(EmailSender sender, EmailSender reply, EmailSender reply1, Contact contact, DateTime waitAfterSending, DateTime currentTime)
		{
			var meetingReplyType = MeetingReplyType.Accept;
			var meetingReplyValues = Enum.GetValues(typeof(MeetingReplyType));
			var rnd = new Random();
			var recuring = Convert.ToBoolean(rnd.Next(2));
			var timeString = currentTime.ToString();
			sender.Subject = "Meeting " + timeString;
			sender.Body = "Test meeting scheduled at: " + timeString;
			sender.StartTime = Helper.GetDateForScheduling(currentTime, rnd);
			sender.Duration = rnd.Next(1, 5) / 2;
			sender.Location = "Conf Room First Floor";
			StringList requiredAttendees = new StringList();
			requiredAttendees.Add(contact.EmailAddress);
			sender.RequiredAttendees = requiredAttendees;

			StringList optionalAttendees = new StringList();
			optionalAttendees.Add(contact.EmailAddress);
			sender.OptionalAttendees = optionalAttendees;

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
			var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
			//var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
			while (waitAfterReply < currentTime)
			{
				currentTime = DateTime.Now;
			}


			Console.WriteLine("Meeting has been scheduled");
		}

		public void CreateTask(EmailSender sender, DateTime currentTime)
		{
			var rnd = new Random();
			var recuring = Convert.ToBoolean(rnd.Next(2));
			string taskSubject = "Task subject " + currentTime;
			string taskBody = "This task was created at " + currentTime;
			sender.StartTime = Helper.GetDateForScheduling(currentTime, rnd);
			sender.Duration = rnd.Next(1, 6);
			int reminder = 15;
			recuring = Convert.ToBoolean(rnd.Next(2));
			sender.CreateTask(taskSubject, taskBody, reminder, recuring);

			Console.WriteLine("Task with subject " + taskSubject + " has been created");
		}

		public void DeleteOldMailsFromAllMailboxes(List<Contact> contactsList)
		{
			foreach (var contact in contactsList)
			{
				EmailSender messageReader = new EmailSender(contact.Credentials);
				messageReader.ReadOldMessages(ConfigManager.DeleteWeekOld);
			}
		}
	}
}
