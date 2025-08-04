using System;
using System.Collections.Generic;
using Independentsoft.Pst;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            PstFile file = new PstFile("c:\\testfolder\\Outlook.pst");

            using (file)
            {
                IList<Folder> folders = file.MailboxRoot.GetFolders();

                for (int i = 0; i < folders.Count; i++)
                {
                    Console.WriteLine("Id: " + folders[i].Id);
                    Console.WriteLine("Name: " + folders[i].DisplayName);
                    Console.WriteLine("Type: " + folders[i].ContainerClass);
                    Console.WriteLine("Item count: " + folders[i].ItemCount);
                    Console.WriteLine("--------------------------------------------------------");
                }
            }

            Console.WriteLine("Press ENTER to exit.");
            Console.Read();
        }
    }
}
