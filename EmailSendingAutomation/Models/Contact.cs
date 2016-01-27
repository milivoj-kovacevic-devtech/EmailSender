using Microsoft.Exchange.WebServices.Data;

namespace EmailSender.Models
{
    public class Contact
    {
        public string Username { get; set; }
		public string Password { get; set; }
		public WebCredentials Credentials
        {
            get
            {
                return new WebCredentials(Username, Password);
            }
        }

        public Contact(string username, string password)
        {
            Username = username;
	        Password = password;
        }
    }
}
