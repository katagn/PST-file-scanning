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
                Folder contacts = file.MailboxRoot.GetFolder("Contacts");

                if (contacts != null)
                {
                    IList<Item> items = contacts.GetItems();

                    for (int m = 0; m < items.Count; m++)
                    {
                        if (items[m] is Contact)
                        {
                            Contact contact = (Contact)items[m];

                            Console.WriteLine("Id: " + contact.Id);
                            Console.WriteLine("GivenName: " + contact.GivenName);
                            Console.WriteLine("Email1Address: " + contact.Email1Address);
                            Console.WriteLine("Email1DisplayName: " + contact.Email1DisplayName);
                            Console.WriteLine("BusinessPhone: " + contact.BusinessPhone);
                            Console.WriteLine("BusinessAddress: " + contact.BusinessAddress);
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
