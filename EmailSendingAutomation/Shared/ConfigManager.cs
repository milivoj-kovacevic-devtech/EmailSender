using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using log4net;

namespace EmailSender.Shared
{
    public static class ConfigManager
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof (ConfigManager));

        public static double HoursToWork = Double.Parse(ConfigurationManager.AppSettings["HoursToWork"]);
        public static bool DeleteWeekOld = Boolean.Parse(ConfigurationManager.AppSettings["DeleteWeekOld"]);
        public static string TextAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "attachment.txt";
        public static string BinaryAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "attachment.exe";
        public static string ReplyAttachment = ConfigurationManager.AppSettings["AttachmentPath"] + "replyAttachment.txt";

        public static List<Contact> GetContacts()
        {
            var contactsList = new List<Contact>();
            try
            {
                var emailUsernames = ConfigurationManager.GetSection("EmailUsernames") as NameValueCollection;
                if (emailUsernames == null) throw new Exception("No users in config file.");
                foreach (var userKey in emailUsernames.AllKeys)
                {
                    var userName = emailUsernames.GetValues(userKey).FirstOrDefault();
                    contactsList.Add(new Contact(userName));
                }
            }
            catch (Exception ex)
            {
                Log.Error("Error getting contacts from config file: " + ex.Message);
            }

            return contactsList;
        }
    }
}
