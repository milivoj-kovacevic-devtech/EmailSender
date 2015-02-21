using log4net;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;

namespace EmailSender
{

    public class EmailSender
    {
		private static readonly ILog Log = LogManager.GetLogger(typeof(EmailSender));

        protected ExchangeService Service;

        public string Subject { get; set; }
        public string Body { get; set; }
        public string FromEmailAddress { get; set; }
        public string ToEmailAddress { get; set; }
        public string AttachmentLocation { get; set; }
        public Guid TestUniqueId { get; set; }
        public ExtendedPropertyDefinition ExtendedProperyDef { get; set; }

        public DateTime StartTime { get; set; }
        public int Duration { get; set; }
        public string Location { get; set; }
        public StringList RequiredAttendees { get; set; }
        public StringList OptionalAttendees { get; set; }

        public EmailSender(WebCredentials credentials)
        {
            ServicePointManager.ServerCertificateValidationCallback =
              ((sender, certificate, chain, sslPolicyErrors) => true);

            Service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

			try
			{

				Service.Url = new Uri(ConfigurationManager.AppSettings["ExchangeAPI"]);
				Service.Credentials = credentials;

				Service.TraceEnabled = true;
				Service.TraceFlags = TraceFlags.All;
			}
			catch (ArgumentNullException e)
			{
				Log.Error("Unable to connect to EWS: " + e.ToString());
				throw;
			}
        }

        public ExtendedPropertyDefinition CreateExtendedPropertyDefinition(string extPropertyName)
        {
            return new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, extPropertyName, MapiPropertyType.String);
        }

        /*
         * Classic email message related methods and functions
         */

        public void SendMessage()
        {
			Log.Info("EmailSender.SendMessage() init...");
            ExtendedProperyDef = CreateExtendedPropertyDefinition("EmailMessageId");

            TestUniqueId = Guid.NewGuid();


            var email = new EmailMessage(Service);

            email.ToRecipients.Add(ToEmailAddress);

            email.Subject = Subject;
            email.Body = new MessageBody(Body);

            email.SetExtendedProperty(ExtendedProperyDef, TestUniqueId.ToString());

            if (AttachmentLocation != null)
            {
				Log.Info("EmailSender.SendMessage() - Adding attachment...");
                email.Attachments.AddFileAttachment(AttachmentLocation);
            }

			Log.Debug("EmailSender.SendMessage() - Sending message...");
            email.SendAndSaveCopy();
			Log.Debug("Success!");

            Console.WriteLine("An email with the subject '" + email.Subject + "' has been sent to '" + email.ToRecipients[0] + "' and saved to SendItems folder.");

			Log.Info("EmailSender.SendMessage() end.");
        }

        public void Reply(string subject)
        {
			Log.Info("EmailSender.Reply(string) init...");
            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, FromEmailAddress), new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, subject));
            var view = new ItemView(1);

			try
			{
				FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);

				EmailMessage reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);

				bool replyToAll = true;
				ResponseMessage responseMessage = reply.CreateReply(replyToAll);

				responseMessage.BodyPrefix = Body;

				Log.Debug("EmailSender.Reply(string) - Sending message...");
				responseMessage.SendAndSaveCopy();
				Log.Debug("Success!");
			} // TODO: Change exception type when caught for the first time
			catch (Exception e)
			{
				Log.Error(e);
				throw;
			}

			Log.Info("EmailSender.Reply(string) end.");
        }

        public void Reply(ExtendedPropertyDefinition extPropDef, Guid testUniqueId)
        {
			Log.Info("EmailSender.Reply(ExtendedPropertyDefinition, Guid) init...");

            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(extPropDef, testUniqueId.ToString()));
            var view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);

			try
			{
				EmailMessage reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);

				bool replyToAll = true;
				ResponseMessage responseMessage = reply.CreateReply(replyToAll);

				responseMessage.BodyPrefix = Body;

				Log.Debug("EmailSender.Reply(ExtendedPropertyDefinition, Guid) - Sending message...");
				responseMessage.SendAndSaveCopy();
				Log.Debug("Success!");
			} // TODO: Change exception type when caught for the first time
			catch (Exception e) 
			{
				Log.Error(e);
				throw;
			}

			Log.Info("EmailSender.Reply(ExtendedPropertyDefinition, Guid) end.");
		}

        /*
         *  Calendar items related methods and functions
         */

		public void ScheduleMeeting(bool recuring)
        {
			Log.Info("EmailSender.ScheduleMeeting() init...");
            var meeting = new Appointment(Service);

            meeting.Subject = Subject;
            meeting.Body = new MessageBody(Body);
            meeting.Start = StartTime;
			meeting.End = StartTime.AddHours(Duration);

			if (recuring)
			{
				Log.Info("SheduleMeeting() - Recuring meeting...");
				Recurrence recurrence = new Recurrence.DailyPattern();
				recurrence.StartDate = StartTime;
				recurrence.EndDate = StartTime.AddDays(10);
				meeting.Recurrence = recurrence;
			}

            meeting.Location = Location;

            if (RequiredAttendees != null)
            {
                foreach (string attendee in RequiredAttendees)
                    meeting.RequiredAttendees.Add(attendee);
            }

            if (OptionalAttendees != null)
            {
                foreach (string attendee in OptionalAttendees)
                    meeting.OptionalAttendees.Add(attendee);
            }

			Log.Debug("EmailSender.ScheduleMeeting() - Sending message...");
            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);
			Log.Debug("Success!");

            Console.WriteLine("An appointment with the subject '" + meeting.Subject + "' has been scheduled.");
		
			Log.Info("EmailSender.ScheduleMeeting() end.");
		}

		public void ReplyToMeetingRequest(string subject, MeetingReplyType replyType)
		{
			Log.Info("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) init...");
			//Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject));
			var view = new ItemView(1);

			try
			{
				FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

				Console.WriteLine(findResults.Count() + " item(s) found");

				MeetingRequest meetingReply = MeetingRequest.Bind(Service, findResults.ElementAt(0).Id);

				meetingReply.Body = new MessageBody(Body);

				Log.Debug("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) - Sending message...");

				if (meetingReply.ConflictingMeetingCount > 0)
				{
					meetingReply.Body = new MessageBody("Sorry. I already have a meeting at requested time.");
					Log.Debug("ConflictingMeetingCount > 0 - Decline meeting...");
					meetingReply.Decline(true);
				}
				else
				{
					switch (replyType)
					{
						case MeetingReplyType.Accept:
							Log.Info("Accept meeting request...");
							meetingReply.Accept(true);
							break;
						case MeetingReplyType.AcceptTentatively:
							Log.Info("Accept tentatively meeting request...");
							meetingReply.AcceptTentatively(true);
							break;
						case MeetingReplyType.Decline:
							Log.Info("Decline meeting request...");
							meetingReply.Decline(true);
							break;
						default:
							Console.WriteLine("This is unknown reply type");
							break;
					}
				}
				Log.Debug("Success!");
			} // TODO: Change exception type when caught for the first time 
			catch (Exception e)
			{
				Log.Error(e);
				throw;
			}

			Log.Info("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) end.");
		}

        /*
         * Task related methods and functions
         */

        public void CreateTask(string subject, string body, int reminder, bool recuring)
        {
			Log.Info("EmailSender.CreateTask(string, string, int, bool) init...");

            // Create the task item and set property values.
            var task = new Task(Service);
            task.Subject = subject;
            task.Body = new MessageBody(body);
            task.StartDate = StartTime;
            task.DueDate = StartTime.AddDays(Duration);
            task.ReminderMinutesBeforeStart = reminder;
            if (recuring)
            {
                task.Recurrence = new Recurrence.DailyPattern(DateTime.Now.AddMinutes(10), 1);
            }

            task.Save();

			Log.Info("EmailSender.CreateTask(string, string, int, bool) end.");
        }

		public void ReadOldMessages(bool deleteWeekOld)
		{
			Log.Info("EmailSender.ReadOldMessages() init...");

			Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
			var view = new ItemView(10);

			FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

			Console.WriteLine(findResults.Count() + " unread item(s) found");

			foreach (Item result in findResults)
			{
				EmailMessage messagge = EmailMessage.Bind(Service, result.Id);
				messagge.IsRead = true;
				messagge.Update(ConflictResolutionMode.AlwaysOverwrite);
				Console.WriteLine("Message with subject: " + messagge.Subject + " has been read!");
			}

			if(deleteWeekOld)
			{
				DeleteOldMessages();
			}

			Log.Info("EmailSender.ReadOldMessages() end.");
		}

		private void DeleteOldMessages()
		{
			Log.Info("EmailSender.DeleteOldMessages() init...");

			var folders = new List<WellKnownFolderName> {WellKnownFolderName.Inbox, WellKnownFolderName.Calendar, WellKnownFolderName.DeletedItems, 
					WellKnownFolderName.Drafts, WellKnownFolderName.SentItems, WellKnownFolderName.Tasks, WellKnownFolderName.JunkEmail};

			foreach (WellKnownFolderName folder in folders)
			{
				DeleteItemsInFolder(folder);
			}
			Log.Info("EmailSender.DeleteOldMessages() end.");
		}

		// Helper method for deleting items older than one week in specified folder
		private void DeleteItemsInFolder(WellKnownFolderName folderName)
		{
			Log.Info("EmailSender.DeleteItemsInFolder() init...");

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeCreated, (DateTime.Now.AddDays(-7))));
			var view = new ItemView(50);

			FindItemsResults<Item> findResults = Service.FindItems(folderName, searchFilter, view);

			foreach (Item result in findResults)
			{
				if (folderName == WellKnownFolderName.Calendar)
				{
					Log.Debug("Deleting meeting...");
					var meeting = Appointment.Bind(Service, result.Id);
					meeting.Delete(DeleteMode.HardDelete, SendCancellationsMode.SendToNone);
				}
				else
				{
					Log.Debug("Deleting item...");
					Item item = Item.Bind(Service, result.Id);
					item.Delete(DeleteMode.HardDelete);
				}
			}

			Log.Info("EmailSender.DeleteItemsInFolder() end.");
		}

    }
}
