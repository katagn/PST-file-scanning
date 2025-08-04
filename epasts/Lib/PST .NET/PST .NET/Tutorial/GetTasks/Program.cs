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
                Folder tasksFolder = file.MailboxRoot.GetFolder("Tasks");

                if (tasksFolder != null)
                {
                    IList<Item> items = tasksFolder.GetItems();

                    for (int m = 0; m < items.Count; m++)
                    {
                        if (items[m] is Task)
                        {
                            Task task = (Task)items[m];

                            Console.WriteLine("Id: " + task.Id);
                            Console.WriteLine("Subject: " + task.Subject);
                            Console.WriteLine("StartDate: " + task.StartDate);
                            Console.WriteLine("DueDate: " + task.DueDate);
                            Console.WriteLine("Owner: " + task.Owner);
                            Console.WriteLine("Body: " + task.Body);
                            Console.WriteLine("----------------------------------------------------------------");
                        }
                    }
                }
            }

            Console.WriteLine("Press ENTER to exit.");
            Console.Read();
        }
    }
}
