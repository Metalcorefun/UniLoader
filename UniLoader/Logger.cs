using System;
using System.IO;

namespace UniLoader
{
    public static class Logger
    {
        private static StreamWriter streamWriter;

        public static void SetFile(string path)
        {
            streamWriter = new StreamWriter(path, false);
        }

        public static void WriteLine(string message)
        {
            streamWriter.WriteLine($"{DateTime.Now:dd.MM.yy hh:mm:ss} - {message}");
            streamWriter.Flush();
        }
    }
}
