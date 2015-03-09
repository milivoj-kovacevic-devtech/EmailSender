using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSender
{
	public static class ConfigManager
	{
		public static double HoursToWork = Double.Parse(ConfigurationManager.AppSettings["HoursToWork"]);
		public static bool DeleteWeekOld = Boolean.Parse(ConfigurationManager.AppSettings["DeleteWeekOld"]);
		public static string TextAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "attachment.txt";
		public static string BinaryAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "attachment.exe";
		public static string ReplyAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "replyAttachment.txt";

		public static List<Contact> GetContacts()
		{
			var contactsList = new List<Contact>(); 
			var emailUsernames = ConfigurationManager.GetSection("EmailUsernames") as NameValueCollection;
			if (emailUsernames != null)
			{
				foreach (var userKey in emailUsernames.AllKeys)
				{
					string userName = emailUsernames.GetValues(userKey).FirstOrDefault();
					contactsList.Add(new Contact(userName));
				}
			}

			return contactsList;
		}
	}
}
