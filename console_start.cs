using System;

namespace Senselenium
{
    class Senselenium
    {
        public static void Main(string[] args)
        {
            if (args.Length >= 1 && args[0].Length > 0)
            {
                MainWindow mw = new MainWindow(args[0]);
            }
            else
            {
                MainWindow mw = new MainWindow("");
            }
            Console.ReadLine();
        }
    }
}
