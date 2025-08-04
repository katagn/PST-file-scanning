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
                IList<Folder> folders = file.MailboxRoot.GetFolders(true);

                string parentFolder = file.MailboxRoot.DisplayName;

                Dictionary<long, string> parents = new Dictionary<long, string>();

                parents.Add(file.MailboxRoot.Id, parentFolder);

                for (int i = 0; i < folders.Count; i++)
                {
                    Folder currentFolder = folders[i];

                    parentFolder = (string)parents[currentFolder.ParentId];

                    string currentFolderPath = parentFolder + "\\" + currentFolder.DisplayName;
                    parents.Add(currentFolder.Id, currentFolderPath);

                    Console.WriteLine("Id: " + currentFolder.Id);
                    Console.WriteLine("Name: " + currentFolder.DisplayName);
                    Console.WriteLine("Type: " + currentFolder.ContainerClass);
                    Console.WriteLine("Item count: " + currentFolder.ItemCount);
                    Console.WriteLine("Path: " + currentFolderPath);
                    Console.WriteLine("----------------------------------------------------------------");
                }
            }

            Console.WriteLine("Press ENTER to exit.");
            Console.Read();
        }
    }
}
