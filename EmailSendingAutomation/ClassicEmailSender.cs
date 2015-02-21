using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace EmailSendingAutomation
{
    public class ClassicEmailSender : EmailSender
    {
        public ClassicEmailSender(WebCredentials credentials)
            : base(credentials)
        {
        }

        public void SendMessage()
        {
            ExtendedProperyDef = CreateExtendedPropertyDefinition("EmailMessageId");

            TestUniqueId = Guid.NewGuid();


            EmailMessage email = new EmailMessage(Service);

            email.ToRecipients.Add(ToEmailAddress);

            email.Subject = Subject;
            email.Body = new MessageBody(Body);

            email.SetExtendedProperty(ExtendedProperyDef, TestUniqueId.ToString());

            email.SendAndSaveCopy();

            Console.WriteLine("An email with the subject '" + email.Subject + "' has been sent to '" + email.ToRecipients[0] + "' and saved to SendItems folder.");
        }

        public void Reply(string subject) 
        {
            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, FromEmailAddress), new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, subject));
            ItemView view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);

            EmailMessage reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);

            bool replyToAll = true;
            ResponseMessage responseMessage = reply.CreateReply(replyToAll);

            responseMessage.BodyPrefix = Body;

            responseMessage.SendAndSaveCopy();
        }

        public void Reply(ExtendedPropertyDefinition extPropDef, Guid testUniqueId)
        {
            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(extPropDef, testUniqueId.ToString()));
            ItemView view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);

            EmailMessage reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);

            bool replyToAll = true;
            ResponseMessage responseMessage = reply.CreateReply(replyToAll);

            responseMessage.BodyPrefix = Body;

            responseMessage.SendAndSaveCopy();
        }


    }
}
