using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EmailSendingAutomation
{

    class Program
    {       
        //static void Main(string[] args)
        //{
        //    DateTime currentTime = DateTime.Now;
        //    DateTime timeToStop = DateTime.Now.AddMinutes(3);
        //    // TODO Should be inside list/array
        //    Contact fedor = new Contact("fedor.hajdu");
        //    Contact nemanja = new Contact("nemanja.tomic");
        //    Contact paula = new Contact("paula.novokmet");

        //    List<Contact> contactsList = new List<Contact>();
        //    contactsList.Add(fedor);
        //    contactsList.Add(nemanja);
        //    contactsList.Add(paula);

        //    Contact[] contactsArray = new Contact[3];
        //    contactsArray[0] = fedor;
        //    contactsArray[1] = nemanja;
        //    contactsArray[2] = paula;

        //    while (currentTime < timeToStop)
        //    {
        //        DateTime waitAfterSending = DateTime.Now.AddMinutes(1);

        //        // should be "randomly" chosen
        //        WebCredentials senderCredentials = fedor.Credentials;

        //        EmailSender sender = new EmailSender(senderCredentials);
        //        sender.Subject = "TestEmail" + currentTime.ToString();
        //        sender.Body = "Test email sent using EWS Managed API";
        //        sender.ToEmailAddress = paula.EmailAddress;
        //        sender.SendMessage();

        //        while (waitAfterSending > currentTime)
        //        {
        //            currentTime = DateTime.Now;
        //        }
               
        //        //WebCredentials recieverCredentials = nemanja.Credentials;
        //        WebCredentials recieverCredentials = paula.Credentials;

        //        EmailSender reply = new EmailSender(recieverCredentials);
        //        reply.Body = "This is reply to test message sent using EWS Managed API" + " extended property included";
        //        reply.FromEmailAddress = fedor.EmailAddress;
        //        reply.Reply(sender.ExtendedProperyDef, sender.TestUniqueId);

        //        DateTime waitAfterReply = DateTime.Now.AddSeconds(15);
        //        while (waitAfterReply < currentTime)
        //        {
        //            currentTime = DateTime.Now;
        //        }

        //    //EmailSender meetingSender = new EmailSender(senderCredentials);
        //    //meetingSender.Subject = "TestMeeting" + currentTime.ToString();
        //    //meetingSender.Body = "Test meeting scheduled using EWS Managed API";
        //    //meetingSender.StartTime = new DateTime(2014, 7, 19, 9, 0, 0);
        //    //meetingSender.Duration = 1;
        //    //meetingSender.Location = "Conf Room First Floor";
        //    //StringList requiredAttendees = new StringList();
        //    //requiredAttendees.Add(nemanja.EmailAddress);
        //    //meetingSender.RequiredAttendees = requiredAttendees;

        //    //StringList optionalAttendees = new StringList();
        //    //optionalAttendees.Add(paula.EmailAddress);
        //    //meetingSender.OptionalAttendees = optionalAttendees;
        //    //meetingSender.ScheduleMeeting();

        //    //string taskSubject = "This is task subject";
        //    //string taskBody = "This needs to be done ASAP";
        //    //meetingSender.StartTime = new DateTime(2014, 7, 19, 10, 30, 0);
        //    //int reminder = 5;
        //    //bool recuring = false;
        //    //meetingSender.CreateTask(taskSubject, taskBody, reminder, recuring);

        //    //waitToReply.StartTimer();

        //    //EmailSender meetingReplySender = new EmailSender(recieverCredentials);
        //    //meetingReplySender.Body = "Reply to test meeting scheduled using EWS Managed API";
        //    //meetingReplySender.AcceptMeeting(meetingSender.Subject);

        //    //EmailSender meetingReplySender1 = new EmailSender(recieverCredentials1);
        //    //meetingReplySender1.Body = "Reply to test meeting scheduled using EWS Managed API";
        //    //meetingReplySender1.DeclineMeeting(meetingSender.Subject);

        //    //    waitAfterReply.StartTimer();

        //    /*
        //    if (emailMessage)
        //    {

        //    } else if (calendarItem)
        //    {

        //    } else
        //    {
        //        Console.WriteLine("Are you kidding me? This is not valid email item.");
        //    }
        //    */

        //        currentTime = DateTime.Now;
        //    }
        //}
    }
}
