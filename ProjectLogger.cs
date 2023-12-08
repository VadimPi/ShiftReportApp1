using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using NLog;

namespace ShiftReportApp1
{
    public class ProjectLogger
    {
        private readonly static Logger logger = LogManager.LoadConfiguration(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NLog.config")).GetCurrentClassLogger();

        public static void LogDebug(string message)
        {
            logger.Debug(message);
        }

        public static void LogInfo(string message)
        {
            logger.Info(message);
        }

        public static void LogWarning(string message)
        {
            logger.Warn(message);
        }

        public static void LogError(string message)
        {
            logger.Error(message);
        }

        public static void LogException(string message, System.Exception ex)
        {
            logger.Error(ex, message);
        }
    }
}
