using NLog;
using System;

namespace ExchangeTracker
{
    class Program
    {
        static void Main()
        {
            var logger = LogManager.GetCurrentClassLogger();
            try
            {
                using (var watcher = new ExchangeWatcher())
                {
                    Console.ReadLine();                    
                }
            }
            catch (Exception ex)
            {
                logger.ErrorException("fatal error, exiting", ex);
            }
        }
    }
}
