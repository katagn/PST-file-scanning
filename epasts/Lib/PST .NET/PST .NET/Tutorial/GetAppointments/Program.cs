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
                Folder calendar = file.MailboxRoot.GetFolder("Calendar");

                if (calendar != null)
                {
                    IList<Item> items = calendar.GetItems();

                    for (int m = 0; m < items.Count; m++)
                    {
                        if (items[m] is Appointment)
                        {
                            Appointment appointment = (Appointment)items[m];

                            Console.WriteLine("Id: " + appointment.Id);
                            Console.WriteLine("Subject: " + appointment.Subject);
                            Console.WriteLine("StartTime: " + appointment.StartTime);
                            Console.WriteLine("EndTime: " + appointment.EndTime);
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
