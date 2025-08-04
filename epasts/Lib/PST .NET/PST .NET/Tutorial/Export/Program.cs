using System;
using System.IO;
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

                Dictionary<long, string> parents = new Dictionary<long, string>();

                string parentFolderPath = "c:\\temp";
                string currentFolderPath = parentFolderPath + "\\" + file.MailboxRoot.DisplayName;

                Directory.CreateDirectory(currentFolderPath);
                parents.Add(file.MailboxRoot.Id, currentFolderPath);

                //Create folder structure
                for (int i = 0; i < folders.Count; i++)
                {
                    Folder currentFolder = folders[i];

                    parentFolderPath = (string)parents[currentFolder.ParentId];
                    currentFolderPath = parentFolderPath + "\\" + currentFolder.DisplayName;

                    Directory.CreateDirectory(currentFolderPath);
                    parents.Add(currentFolder.Id, currentFolderPath);
                }

                //Get items and save to folders as .msg files
                for (int j = 0; j < folders.Count; j++)
                {
                    for (int k = 0; k < folders[j].ChildrenCount; k += 100)
                    {
                        IList<Item> items = folders[j].GetItems(k, k + 100);

                        for (int m = 0; m < items.Count; m++)
                        {
                            string parentFolder = (string)parents[items[m].ParentId];
                            string fileName = GetFileName(items[m].Subject);

                            string filePath = parentFolder + "\\" + fileName;

                            if (filePath.Length > 244)
                            {
                                filePath = filePath.Substring(0, 244);
                            }

                            filePath = filePath + "-" + items[m].Id + ".msg";

                            items[m].Save(filePath);
                        }
                    }
                }
            }
        }

        private static string GetFileName(string subject)
        {
            if (subject == null || subject.Length == 0)
            {
                string fileName = "NoSubject";
                return fileName;
            }
            else
            {
                string fileName = "";

                for (int i = 0; i < subject.Length; i++)
                {
                    if (subject[i] > 31 && subject[i] < 127)
                    {
                        fileName += subject[i];
                    }
                }

                fileName = fileName.Replace("\\", "_");
                fileName = fileName.Replace("/", "_");
                fileName = fileName.Replace(":", "_");
                fileName = fileName.Replace("*", "_");
                fileName = fileName.Replace("?", "_");
                fileName = fileName.Replace("\"", "_");
                fileName = fileName.Replace("<", "_");
                fileName = fileName.Replace(">", "_");
                fileName = fileName.Replace("|", "_");

                return fileName;
            }
        }
    }
}
