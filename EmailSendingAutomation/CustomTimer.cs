using System;
using System.Timers;

namespace EmailSendingAutomation
{
    class CustomTimer
    {
        private Timer aTimer;
        public bool TimerEnabled { get { return aTimer.Enabled; } }
        private DateTime EndTime { get; set; }

        public CustomTimer(int minutesDuration)
        {
            EndTime = DateTime.Now.AddMinutes(minutesDuration);
            // Create a timer with a two seconds interval.
            aTimer = new System.Timers.Timer(5000);
            aTimer.Enabled = false;

        }

        public void StartTimer()
        {
            // Hook up the Elapsed event for the timer.
            aTimer.Elapsed += OnTimedEvent;
            aTimer.Start();
        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            if (e.SignalTime > EndTime)
            {
                aTimer.Stop();
            }
        }
    }
}
