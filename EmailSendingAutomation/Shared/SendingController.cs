using System.Linq;
using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;

namespace EmailSender.Shared
{
    public class SendingController
    {
        private static ConfigManager _config;

        public SendingController(ConfigManager config)
        {
            _config = config;
        }

        public void SendEmail(Contact senderContact, Contact replyContact)
		{
			var rnd = new Random();
            var attachments = new[] { _config.GetTextAttachmentPath(), _config.GetBinaryAttachmentPath() };
		    var currentTime = DateTime.Now;
            var waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
            var timeString = currentTime.ToString();
		    var sender = new EmailSender(senderContact.Credentials)
		    {
		        Subject = "Email " + timeString,
		        Body = "This email was sent at: " + timeString,
		        ToEmailAddress = replyContact.EmailAddress
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

            var	reply = new EmailSender(replyContact.Credentials)
		    {
		        Body = "This is reply to test message sent using EWS Managed API. It was sent at " + currentTime,
		        FromEmailAddress = replyContact.EmailAddress
		    };

		    if (rnd.Next(2) == 0)
			{
				reply.AttachmentLocation = _config.GetReplyAttachmentPath();
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

        public void ScheduleMeeting(Contact senderContact, Contact replyContact, Contact reply1Contact)
        {
            var meetingReplyValues = Enum.GetValues(typeof(MeetingReplyType));
            var rnd = new Random();
            var recuring = Convert.ToBoolean(rnd.Next(2));
            var currentTime = DateTime.Now;
            var waitAfterSending = DateTime.Now.AddMinutes(rnd.Next(1, 5)); // Wait to reply to email
            var timeString = currentTime.ToString();
            var sender = new EmailSender(senderContact.Credentials)
            {
                Subject = "Meeting " + timeString,
                Body = "Test meeting scheduled at: " + timeString,
                StartTime = Helper.GetDateForScheduling(currentTime, rnd),
                Duration = rnd.Next(1, 5) / 2,
                Location = "Conf Room First Floor"
            };
            sender.RequiredAttendees.Add(replyContact.EmailAddress);
            sender.OptionalAttendees.Add(reply1Contact.EmailAddress);

            sender.ScheduleMeeting(recuring);

            while (waitAfterSending > currentTime)
            {
                currentTime = DateTime.Now;
                Thread.Sleep(10000);
            }

            timeString = currentTime.ToString();
            var reply = new EmailSender(replyContact.Credentials)
            {
                Body = "Reply from required attendee to test meeting sent at: " + timeString
            };
            var meetingReplyType = (MeetingReplyType)meetingReplyValues.GetValue(rnd.Next(3));
            reply.ReplyToMeetingRequest(sender.Subject, meetingReplyType);

            var reply1 = new EmailSender(reply1Contact.Credentials)
            {
                Body = "Reply from optional attendee to test meeting sent at " + timeString
            };
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

        public void CreateTask(Contact senderContact)
        {
            var rnd = new Random();
            var recuring = Convert.ToBoolean(rnd.Next(2));
            var currentTime = DateTime.Now;
            var taskSubject = "Task subject " + currentTime;
            var taskBody = "This task was created at " + currentTime;
            var sender = new EmailSender(senderContact.Credentials)
            {
                StartTime = Helper.GetDateForScheduling(currentTime, rnd),
                Duration = rnd.Next(1, 6)
            };
            const int reminder = 15;
            sender.CreateTask(taskSubject, taskBody, reminder, recuring);

            Console.WriteLine("Task with subject " + taskSubject + " has been created");
        }

        public void DeleteOldMailsFromAllMailboxes(List<Contact> contactsList)
        {
            foreach (var messageReader in contactsList.Select(contact => new EmailSender(contact.Credentials)))
            {
                messageReader.ReadOldMessages(_config.GetDeleteOldMailFlag());
            }
        }
    }
}
