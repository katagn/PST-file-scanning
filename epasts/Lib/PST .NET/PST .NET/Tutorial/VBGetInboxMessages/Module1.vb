Imports System
Imports System.Collections.Generic
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
                            Console.WriteLine("DisplayTo: " & message.DisplayTo)
                            Console.WriteLine("DisplayCc: " & message.DisplayCc)
                            Console.WriteLine("SenderName: " & message.SenderName)
                            Console.WriteLine("SenderEmailAddress: " & message.SenderEmailAddress)
                            Console.WriteLine("----------------------------------------------------------------")
                        End If
                    Next
                End If
            End Using

            Console.WriteLine("Press ENTER to exit.")
            Console.Read()

        End Sub
    End Class
End Namespace