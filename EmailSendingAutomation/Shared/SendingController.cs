using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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
			//waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(1, 3)); // Wait for sending new email
			var waitAfterReply = DateTime.Now.AddMinutes(rnd.Next(3, 13)); // Wait for sending new email
			while (waitAfterReply < currentTime)
			{
				currentTime = DateTime.Now;
			}

			Console.WriteLine("email message was sent");
		}

		public void ScheduleMeeting()
		{
			throw new NotImplementedException();
		}

		public void CreateTask()
		{
			throw new NotImplementedException();
		}
	}
}
