﻿using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using log4net;

namespace EmailSender.Shared
{
    public class ConfigManager
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof (ConfigManager));

        public bool GetDeleteOldMailFlag()
        {
            return GetBooleanConfigValue("DeleteWeekOld");
        }

        public string GetTextAttachmentPath()
        {
            return GetAttachmentsFolderPath() + "attachment.txt";
        }
        
        public string GetBinaryAttachmentPath()
        {
            return GetAttachmentsFolderPath() + "attachment.exe";
        }

        public string GetReplyAttachmentPath()
        {
            return GetAttachmentsFolderPath() + "replyAttachment.txt";
        }

        private string GetAttachmentsFolderPath()
        {
            return GetStringConfigValue("AttachmentPath");
        }

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
