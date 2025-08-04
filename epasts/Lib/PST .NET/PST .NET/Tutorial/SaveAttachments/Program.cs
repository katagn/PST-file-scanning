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
                Folder inbox = file.MailboxRoot.GetFolder("Inbox");

                if (inbox != null)
                {
                    IList<Item> items = inbox.GetItems();

                    for (int m = 0; m < items.Count; m++)
                    {
                        for (int i = 0; i < items[m].Attachments.Count; i++)
                        {
                            Attachment attachment = items[m].Attachments[i];

                            string fileName = (attachment.FileName != null) ? attachment.FileName : attachment.DisplayName;

                            string filePath = "c:\\temp\\" + fileName;

                            attachment.Save(filePath, true);
                        }
                    }
                }
            }
        }
    }
}
