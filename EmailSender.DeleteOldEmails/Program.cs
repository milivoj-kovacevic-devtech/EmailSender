using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmailSender;

namespace EmailSender.DeleteOldEmails
{
	class Program
	{
		static void Main(string[] args)
		{
			Contact paula = new Contact("paula.novokmet");

			DateTime current = DateTime.Now;

			EmailSender mailbox = new EmailSender(paula.Credentials);

			mailbox.ReadOldMessages();
		}
	}
}
