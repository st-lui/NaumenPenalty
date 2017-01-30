using System;
using NLog;
using NLog.Config;
using NLog.Targets;

public class MyLogger
{
	private static Logger logger;
	public static Logger GetInstance()
	{
		if (logger == null)
			logger = createLogger();
		return logger;
	}

	private static Logger createLogger()
	{
		LoggingConfiguration config = new LoggingConfiguration();

		FileTarget fileTarget = new FileTarget();
		fileTarget.FileName= "C:\\projects\\logs\\np\\log.txt";
		fileTarget.ArchiveEvery = FileArchivePeriod.Day;
		fileTarget.MaxArchiveFiles = 31;
		fileTarget.ArchiveNumbering=ArchiveNumberingMode.Date;
		fileTarget.Layout="${date}|${level}|${message}";
		config.LoggingRules.Add(new LoggingRule("SingletonLogger", LogLevel.Trace, fileTarget));
		LogManager.Configuration = config;
		return LogManager.GetLogger("SingletonLogger");
	}
}