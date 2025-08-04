using System;
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
                Console.WriteLine("Message store name: " + file.MessageStore.DisplayName);
                Console.WriteLine("Mailbox root folder: " + file.MailboxRoot.DisplayName);
                Console.WriteLine("Encryption type: " + file.EncryptionType);
                Console.WriteLine("File size: " + file.Size);
                Console.WriteLine("Is 64-bit: " + file.Is64Bit);
            }

            Console.WriteLine("Press ENTER to exit.");
            Console.Read();
        }
    }
}
