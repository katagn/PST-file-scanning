Imports System
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Dim inbox As Folder = file.MailboxRoot.GetFolder("Inbox")

                If inbox IsNot Nothing Then

                    Dim items As IList(Of Item) = inbox.GetItems()

                    For m As Integer = 0 To items.Count - 1

                        If TypeOf items(m) Is Message Then

                            Dim message As Message = DirectCast(items(m), Message)

                            Console.WriteLine("Id: " & message.Id)
                            Console.WriteLine("Subject: " & message.Subject)

                            For r As Integer = 0 To message.Recipients.Count - 1
                                Dim recipient As Recipient = message.Recipients(r)

                                Console.WriteLine("Name: " & recipient.DisplayName)
                                Console.WriteLine("Email address: " & recipient.EmailAddress)
                                Console.WriteLine("Recipient type: " & recipient.RecipientType.ToString())
                            Next

                            Console.WriteLine("-------------------------------------------------------")
                        End If
                    Next
                End If
            End Using

            Console.WriteLine("Press ENTER to exit.")
            Console.Read()

        End Sub
    End Class
End Namespace