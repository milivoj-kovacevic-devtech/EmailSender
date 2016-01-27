using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;

namespace JournalItemsCreator.Shared
{
    public class ConfigManager
    {
        public List<string> GetMailboxes()
        {
            var contactsList = new List<string>();
            try
            {
                var mailboxes = ConfigurationManager.GetSection("Mailboxes") as NameValueCollection;
                if (mailboxes == null) throw new Exception("No users in config file.");
                foreach (var userKey in mailboxes.AllKeys)
                {
                    var address = mailboxes.GetValues(userKey).FirstOrDefault();
                    contactsList.Add(address);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error getting contacts from config file: " + ex.Message);
                Console.Read();
            }

            return contactsList;
        }

        public string GetLogFilePath()
        {
            return GetStringConfigValue("LogFilePath");
        }

        public string GetEwsUrl()
        {
            return GetStringConfigValue("EwsUrl");
        }

        internal string GetMailboxPassword()
        {
            return GetStringConfigValue("MailboxPassword");
        }

        protected virtual object ReadConfigValue(string configValue)
        {
            object returnValue = null;

            try
            {
                returnValue = ConfigurationManager.AppSettings[configValue];
            }
            catch
            {
                // ignored
            }

            return returnValue;
        }

        protected string GetStringConfigValue(string configString)
        {
            var returnValue = string.Empty;

            object configValue = ReadConfigValue(configString);
            if (configValue != null)
                returnValue = configValue.ToString();

            return returnValue;
        }

        protected bool GetBooleanConfigValue(string configString)
        {
            var returnValue = false;

            object configValue = ReadConfigValue(configString);
            if (configValue != null)
                returnValue = (configValue.ToString().ToLower() == "true");

            return returnValue;
        }

        protected int GetIntegerConfigValue(string configString)
        {
            var liReturnValue = 0;

            try
            {
                liReturnValue = Convert.ToInt32(ReadConfigValue(configString));
            }
            catch
            {
                // do nothing
            }

            return liReturnValue;
        }
    }
}
