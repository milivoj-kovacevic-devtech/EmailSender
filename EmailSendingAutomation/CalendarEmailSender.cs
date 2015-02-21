using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmailSendingAutomation
{
    class CalendarEmailSender : EmailSender
    {

        public DateTime StartTime { get; set; }
        public int Duration { get; set; }
        public string Location { get; set; }
        public StringList RequiredAttendees { get; set; }
        public StringList OptionalAttendees { get; set; }

        public CalendarEmailSender(string username, string password)
            : base(new WebCredentials(username, password))
        {
        }

        public void ScheduleMeeting()
        {
            Appointment meeting = new Appointment(Service);

            meeting.Subject = Subject;
            meeting.Body = new MessageBody(Body);
            meeting.Start = StartTime;
            meeting.End = StartTime.AddHours(Duration);

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

            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);

            Console.WriteLine("An appointment with the subject '" + meeting.Subject + "' has been scheduled.");
        }

        public void AcceptMeeting(string subject)
        {
            //Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject));
            ItemView view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

            Console.WriteLine(findResults.Count() + " item(s) found");

            MeetingRequest meetingReply = MeetingRequest.Bind(Service, findResults.ElementAt(0).Id);

            meetingReply.Accept(true);

            //bool replyToAll = false;
            //ResponseMessage acceptMessage = meetingReply.CreateReply(replyToAll);

            //acceptMessage.Body = new MessageBody("Reply to meeting invitation");
            //acceptMessage.Sensitivity = Sensitivity.Private;
            //acceptMessage.SendAndSaveCopy();
        }

        public void AcceptMeetingTentatively(string subject)
        {
            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject));
            ItemView view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

            Console.WriteLine(findResults.Count() + " item(s) found");

            MeetingRequest meetingReply = MeetingRequest.Bind(Service, findResults.ElementAt(0).Id);

            meetingReply.AcceptTentatively(true);
        }

        public void DeclineMeeting(string subject)
        {
            Folder inbox = Folder.Bind(Service, WellKnownFolderName.Inbox);

            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject));
            ItemView view = new ItemView(1);

            FindItemsResults<Item> findResults = Service.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

            Console.WriteLine(findResults.Count() + " item(s) found");

            MeetingRequest meetingReply = MeetingRequest.Bind(Service, findResults.ElementAt(0).Id);

            meetingReply.Decline(true);
        }

        public void CreateTask(string subject, string body, DateTime startDate, DateTime dueDate, int reminder)
        {
            // Create the task item and set property values.
            Task task = new Task(Service);
            task.Subject = subject;
            task.Body = new MessageBody(body);
            task.StartDate = startDate;
            task.DueDate = dueDate;
            task.Recurrence = new Recurrence.DailyPattern(DateTime.Now.AddMinutes(10), 1);
            task.ReminderMinutesBeforeStart = reminder;

            task.Save();
        }

    }
}
