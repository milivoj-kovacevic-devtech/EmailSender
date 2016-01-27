using System;
using System.IO;

namespace EmailSender.Shared
{
	public class Logger
	{
		private readonly string _logFilePath;

		public Logger(ConfigManager config)
		{
			_logFilePath = config.GetLogFilePath();
		}

		public void Info(string message)
		{
			using (var stream = GetFilestream())
			{
				stream.WriteLine("[{0}] - {1}", DateTime.Now, message);
			}
			Console.WriteLine(message);
		}

		public void Debug(string message)
		{
			using (var stream = GetFilestream())
			{
				stream.WriteLine("[{0}] - DEBUG: {1}", DateTime.Now, message);
			}
			Console.WriteLine(message);
		}

		public void Error(string message, Exception ex)
		{
			using (var stream = GetFilestream())
			{
				stream.WriteLine("[{0}] - ERROR: {1}", DateTime.Now, message);
			}
			Console.WriteLine(message);
		}

		public void Error(string message)
		{
			using (var stream = GetFilestream())
			{
				stream.WriteLine("[{0}] - ERROR: {1}", DateTime.Now, message);
			}
			Console.WriteLine(message);
		}

		private StreamWriter GetFilestream()
		{
			if (!File.Exists(_logFilePath))
			{
				File.Create(_logFilePath).Dispose();
			}
			return File.AppendText(_logFilePath);
		}
	}
}
