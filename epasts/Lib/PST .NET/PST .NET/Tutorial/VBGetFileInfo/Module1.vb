Imports System
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Console.WriteLine("Message store name: " & file.MessageStore.DisplayName)
                Console.WriteLine("Mailbox root folder: " & file.MailboxRoot.DisplayName)
                Console.WriteLine("Encryption type: " & file.EncryptionType)
                Console.WriteLine("File size: " & file.Size)
                Console.WriteLine("Is 64-bit: " & file.Is64Bit)

            End Using

            Console.WriteLine("Press ENTER to exit.")
            Console.Read()

        End Sub
    End Class
End Namespace