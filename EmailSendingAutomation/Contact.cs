using Microsoft.Exchange.WebServices.Data;
using System.Configuration;

namespace EmailSender
{
    public class Contact
    {
        private readonly string _password = ConfigurationManager.AppSettings["ExchangePassword"];
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
                return new WebCredentials(Username, _password);
            }
        }

        public Contact(string username)
        {
            Username = username;
        }
    }
}
