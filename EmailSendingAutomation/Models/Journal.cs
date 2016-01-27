using System;

namespace EmailSender.Models
{
	public class Journal
	{
		public string Subject { get; set; }
		public string Body { get; set; }
		public string Type { get; set; }
		public string TypeDescription { get; set; }
		public DateTime StartTime { get; set; }
		public DateTime EndTime { get; set; }
		public string[] Companies { get; set; }
		public int IconIndex { get; set; }
	}
}
