using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSender
{
    public class Contact
    {
        private string Password = ConfigurationManager.AppSettings["ExchangePassword"];
        public string Username { get; set; }
        public string EmailAddress 
        { 
            get
            {
                return Username + "@litwareinc.com";
            }
        }

        public WebCredentials Credentials
        {
            get
            {
                return new WebCredentials(Username, Password);
            }
        }

        public Contact(string username)
        {
            Username = username;
        }
    }
}
