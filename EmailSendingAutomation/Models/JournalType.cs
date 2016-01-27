namespace EmailSender.Models
{
	public class JournalType
	{
		public string Type { get; set; }
		public int IconIndex { get; set; }

		public JournalType(string type, int iconIndex)
		{
			Type = type;
			IconIndex = iconIndex;
		}
	}
}
