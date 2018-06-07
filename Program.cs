using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using WUApiLib;

namespace ConsoleApplication1WindowsUpdate
{   
    
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0 || args.Length > 1)
            {
                Console.WriteLine("USAGE:   /? -- this info");
                Console.WriteLine("         /detail or /d -- outputs details to installed.txt and available.txt");
                Console.WriteLine("         /console or /c -- outputs counts to console");
            }
            if (args.Length == 1)
            {
                string a = args[0].ToUpper();
                if (a == "DETAIL" || a == "/DETAIL" || a == "/D")
                {
                    InstalledUpdates(1);
                    UpdatesAvailable(1);
                }

                if (a == "CONSOLE" || a == "/CONSOLE" || a == "/C")
                {
                    InstalledUpdates(0);
                    UpdatesAvailable(0);
                }

                if (a == "?" || a == "/?")
                {
                    Console.WriteLine("USAGE:   /? -- this info");
                    Console.WriteLine("         /detail or /d -- outputs details to installed.txt and available.txt");
                    Console.WriteLine("         /console or /c -- outputs counts to console");
                }                
            }
        }

            public static void InstalledUpdates(int x)
        {
            UpdateSession updateSession = new UpdateSession();
            IUpdateSearcher updateSearchResult = updateSession.CreateUpdateSearcher();
            updateSearchResult.Online = true;//checks for updates online
            ISearchResult searchResults = updateSearchResult.Search("IsInstalled=1 AND IsHidden=0");
            //for the above search criteria refer to 
            //http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=VS.85).aspx
            //Check the remarks section
            if (x == 0)
            {
                Console.WriteLine("Number of updates installed: " + searchResults.Updates.Count);
            }
            if (x == 1)
            {
                if(File.Exists("installed.txt"))
                {
                    File.Delete("installed.txt");
                }
                string summary = "Installed update count: " + searchResults.Updates.Count;
                File.AppendAllText("installed.txt", summary + Environment.NewLine);
                foreach (IUpdate z in searchResults.Updates)
                {
                   string content = z.Title;
                   File.AppendAllText("installed.txt",content + Environment.NewLine);
                }
                Console.WriteLine("Installed Update details written to installed.txt");
            }
        }
        public static void UpdatesAvailable(int y)
            {
                UpdateSession updateSession = new UpdateSession();
                IUpdateSearcher updateSearchResult = updateSession.CreateUpdateSearcher();
                updateSearchResult.Online = true;//checks for updates online
                ISearchResult searchResults = updateSearchResult.Search("IsInstalled=0 AND IsPresent=0");
                //for the above search criteria refer to 
                //http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=VS.85).aspx
                //Check the remarks section
                if (y == 0)
                {
                    Console.WriteLine("Number of updates available: " + searchResults.Updates.Count);
                }
                if (y == 1)
                {
                    if (File.Exists("available.txt"))
                    {
                        File.Delete("available.txt");
                    }
                    string summary = "Available update count: " + searchResults.Updates.Count;
                    File.AppendAllText("available.txt", summary + Environment.NewLine);
                    foreach (IUpdate z in searchResults.Updates)
                    {
                        string content = z.Title;
                        File.AppendAllText("available.txt", content + Environment.NewLine);
                    }
                    Console.WriteLine("Available Update details written to available.txt");
                }
            }
    }
}
