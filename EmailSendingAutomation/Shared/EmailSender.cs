using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using EmailSender.Models;
using Microsoft.Exchange.WebServices.Data;

namespace EmailSender.Shared
{

    public class EmailSender
    {
		private static readonly ConfigManager Config = new ConfigManager();
	    private static Logger _log;

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

        public EmailSender(WebCredentials credentials, Logger log)
        {
	        _log = log;
            ServicePointManager.ServerCertificateValidationCallback =
              ((sender, certificate, chain, sslPolicyErrors) => true);

            Service = new ExchangeService(ExchangeVersion.Exchange2010);

			try
			{

				Service.Url = new Uri(Config.GetExchangeApiUrl());
				Service.Credentials = credentials;

				Service.TraceEnabled = true;
				Service.TraceFlags = TraceFlags.All;
			}
			catch (ArgumentNullException ex)
			{
				_log.Error("Unable to connect to EWS: " + ex.Message);
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
			_log.Info("EmailSender.SendMessage() init...");
			ExtendedProperyDef = CreateExtendedPropertyDefinition("EmailMessageId");

            TestUniqueId = Guid.NewGuid();


            var email = new EmailMessage(Service);

            email.ToRecipients.Add(ToEmailAddress);

            email.Subject = Subject;
            email.Body = new MessageBody(Body);

            email.SetExtendedProperty(ExtendedProperyDef, TestUniqueId.ToString());

            if (AttachmentLocation != null)
            {
				_log.Info("EmailSender.SendMessage() - Adding attachment...");
				email.Attachments.AddFileAttachment(AttachmentLocation);
            }

			_log.Debug("EmailSender.SendMessage() - Sending message...");
			email.SendAndSaveCopy();
			_log.Debug("Success!");

			Console.WriteLine("An email with the subject '" + email.Subject + "' has been sent to '" + email.ToRecipients[0] + "' and saved to SendItems folder.");

			_log.Info("EmailSender.SendMessage() end.");
		}

        public void Reply(string subject)
        {
			_log.Info("EmailSender.Reply(string) init...");
			var sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, FromEmailAddress), new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, subject));
            var view = new ItemView(1);

			try
			{
				var findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);
				var reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);
			    var responseMessage = reply.CreateReply(true);
				responseMessage.BodyPrefix = Body;
				_log.Debug("EmailSender.Reply(string) - Sending message...");
				responseMessage.SendAndSaveCopy();
				_log.Debug("Success!");
			}
            catch (Exception ex)
			{
				_log.Error("An error occured while replying: " + ex.Message);
				throw;
			}

			_log.Info("EmailSender.Reply(string) end.");
		}

        public void Reply(ExtendedPropertyDefinition extPropDef, Guid testUniqueId)
        {
			_log.Info("EmailSender.Reply(ExtendedPropertyDefinition, Guid) init...");

			SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(extPropDef, testUniqueId.ToString()));
            var view = new ItemView(1);

            var findResults = Service.FindItems(WellKnownFolderName.Inbox, sf, view);

			try
			{
				var reply = EmailMessage.Bind(Service, findResults.ElementAt(0).Id, BasePropertySet.IdOnly);
			    var responseMessage = reply.CreateReply(true);
				responseMessage.BodyPrefix = Body;
				_log.Debug("EmailSender.Reply(ExtendedPropertyDefinition, Guid) - Sending message...");
				responseMessage.SendAndSaveCopy();
				_log.Debug("Success!");
			}
			catch (Exception ex) 
			{
				_log.Error("An error occured while replying: " + ex.Message);
				throw;
			}

			_log.Info("EmailSender.Reply(ExtendedPropertyDefinition, Guid) end.");
		}

        /*
         *  Calendar items related methods and functions
         */

		public void ScheduleMeeting(bool recuring)
        {
			_log.Info("EmailSender.ScheduleMeeting() init...");
			var meeting = new Appointment(Service)
            {
                Subject = Subject,
                Body = new MessageBody(Body),
                Start = StartTime,
                End = StartTime.AddHours(Duration)
            };

		    if (recuring)
			{
				_log.Info("SheduleMeeting() - Recuring meeting...");
				Recurrence recurrence = new Recurrence.DailyPattern();
				recurrence.StartDate = StartTime;
				recurrence.EndDate = StartTime.AddDays(10);
				meeting.Recurrence = recurrence;
			}

            meeting.Location = Location;

            if (RequiredAttendees != null)
            {
                foreach (var attendee in RequiredAttendees)
                    meeting.RequiredAttendees.Add(attendee);
            }

            if (OptionalAttendees != null)
            {
                foreach (var attendee in OptionalAttendees)
                    meeting.OptionalAttendees.Add(attendee);
            }

			_log.Debug("EmailSender.ScheduleMeeting() - Sending message...");
			meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);
			_log.Debug("Success!");

			Console.WriteLine("An appointment with the subject '" + meeting.Subject + "' has been scheduled.");

			_log.Info("EmailSender.ScheduleMeeting() end.");
		}

		public void ReplyToMeetingRequest(string subject, MeetingReplyType replyType)
		{
			_log.Info("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) init...");

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject));
			var view = new ItemView(1);

			try
			{
				var findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

				Console.WriteLine(findResults.Count() + " item(s) found");

				var meetingReply = MeetingRequest.Bind(Service, findResults.ElementAt(0).Id);
				meetingReply.Body = new MessageBody(Body);
				_log.Debug("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) - Sending message...");

				if (meetingReply.ConflictingMeetingCount > 0)
				{
					meetingReply.Body = new MessageBody("Sorry. I already have a meeting at requested time.");
					_log.Debug("ConflictingMeetingCount > 0 - Decline meeting...");
					meetingReply.Decline(true);
				}
				else
				{
					switch (replyType)
					{
						case MeetingReplyType.Accept:
							_log.Info("Accept meeting request...");
							meetingReply.Accept(true);
							break;
						case MeetingReplyType.AcceptTentatively:
							_log.Info("Accept tentatively meeting request...");
							meetingReply.AcceptTentatively(true);
							break;
						case MeetingReplyType.Decline:
							_log.Info("Decline meeting request...");
							meetingReply.Decline(true);
							break;
						default:
							Console.WriteLine("This is unknown reply type");
							break;
					}
				}
				_log.Debug("Success!");
			}
			catch (Exception ex)
			{
				_log.Error("An error occurred while replying to meeting request: " + ex.Message);
				throw;
			}

			_log.Info("EmailSender.ReplyToMeetingRequest(string, MeetingReplyType) end.");
		}

        /*
         * Task related methods and functions
         */

        public void CreateTask(string subject, string body, int reminder, bool recuring)
        {
			_log.Info("EmailSender.CreateTask(string, string, int, bool) init...");

			// Create the task item and set property values.
			var task = new Task(Service)
            {
                Subject = subject,
                Body = new MessageBody(body),
                StartDate = StartTime,
                DueDate = StartTime.AddDays(Duration),
                ReminderMinutesBeforeStart = reminder
            };
            if (recuring)
            {
                task.Recurrence = new Recurrence.DailyPattern(DateTime.Now.AddMinutes(10), 1);
            }

            task.Save();

			_log.Info("EmailSender.CreateTask(string, string, int, bool) end.");
		}

		public void ReadOldMessages(bool deleteWeekOld)
		{
			_log.Info("EmailSender.ReadOldMessages() init...");

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
			var view = new ItemView(10);

			var findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

			Console.WriteLine(findResults.Count() + " unread item(s) found");

			foreach (var result in findResults)
			{
				var messagge = EmailMessage.Bind(Service, result.Id);
				messagge.IsRead = true;
				messagge.Update(ConflictResolutionMode.AlwaysOverwrite);
				Console.WriteLine("Message with subject: " + messagge.Subject + " has been read!");
			}

			if(deleteWeekOld)
			{
				DeleteOldMessages();
			}

			_log.Info("EmailSender.ReadOldMessages() end.");
		}

		private void DeleteOldMessages()
		{
			_log.Info("EmailSender.DeleteOldMessages() init...");

			var folders = new List<WellKnownFolderName> {WellKnownFolderName.Inbox, WellKnownFolderName.Calendar, WellKnownFolderName.DeletedItems, 
					WellKnownFolderName.Drafts, WellKnownFolderName.SentItems, WellKnownFolderName.Tasks, WellKnownFolderName.JunkEmail};

			foreach (var folder in folders)
			{
				DeleteItemsInFolder(folder);
			}
			_log.Info("EmailSender.DeleteOldMessages() end.");
		}

		// Helper method for deleting items older than one week in specified folder
		private void DeleteItemsInFolder(WellKnownFolderName folderName)
		{
			_log.Info("EmailSender.DeleteItemsInFolder() init...");

			SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsLessThan(EmailMessageSchema.DateTimeCreated, (DateTime.Now.AddDays(-7))));
			var view = new ItemView(50);

			var findResults = Service.FindItems(folderName, searchFilter, view);

			foreach (var result in findResults)
			{
				if (folderName == WellKnownFolderName.Calendar)
				{
					_log.Debug("Deleting meeting...");
					var meeting = Appointment.Bind(Service, result.Id);
					meeting.Delete(DeleteMode.HardDelete, SendCancellationsMode.SendToNone);
				}
				else
				{
					_log.Debug("Deleting item...");
					var item = Item.Bind(Service, result.Id);
					item.Delete(DeleteMode.HardDelete);
				}
			}

			_log.Info("EmailSender.DeleteItemsInFolder() end.");
		}

	    private static Guid JournalExtendedPropertyGuid
		{
			get { return new Guid("0006200A-0000-0000-C000-000000000046"); }
		}

	    private static Guid ContactExtendedPropertyGuid
	    {
		    get
		    {
			    return new Guid("00062008-0000-0000-C000-000000000046");
		    }
	    }

		public void CreateJournalItem(Journal journal)
	    {
			var typeProp = new ExtendedPropertyDefinition(JournalExtendedPropertyGuid, 34560, MapiPropertyType.String);
		    var typePropDesc = new ExtendedPropertyDefinition(JournalExtendedPropertyGuid, 34578, MapiPropertyType.String);
			var startTimeProp = new ExtendedPropertyDefinition(JournalExtendedPropertyGuid, 34566, MapiPropertyType.SystemTime);
			var endTimeProp = new ExtendedPropertyDefinition(JournalExtendedPropertyGuid, 34568, MapiPropertyType.SystemTime);
			var companiesProp = new ExtendedPropertyDefinition(ContactExtendedPropertyGuid, 34105, MapiPropertyType.StringArray);
			var iconIndexProp = new ExtendedPropertyDefinition(4224, MapiPropertyType.Integer);

			var newJournal = new EmailMessage(Service);
			newJournal.ItemClass = "IPM.Activity";
			newJournal.Subject = journal.Subject;
			newJournal.Body = journal.Body;
			newJournal.SetExtendedProperty(typeProp, journal.Type);
			newJournal.SetExtendedProperty(typePropDesc, journal.TypeDescription);
			newJournal.SetExtendedProperty(startTimeProp, journal.StartTime);
			newJournal.SetExtendedProperty(endTimeProp, journal.EndTime);
			newJournal.SetExtendedProperty(iconIndexProp, journal.IconIndex);
			newJournal.SetExtendedProperty(companiesProp, journal.Companies);

			newJournal.Save(WellKnownFolderName.Journal);
		}
	}
}
