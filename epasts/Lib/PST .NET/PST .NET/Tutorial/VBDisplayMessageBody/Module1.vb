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
                        Console.WriteLine("Id: " & items(m).Id)
                        Console.WriteLine("Subject: " & items(m).Subject)

                        Console.WriteLine("Plain body:")
                        Console.WriteLine(items(m).Body)
                        Console.WriteLine("-------------------------------------------------------")

                        Console.WriteLine("Html body:")
                        Console.WriteLine(items(m).BodyHtmlText)
                        Console.WriteLine("-------------------------------------------------------")

                        If items(m).BodyRtf IsNot Nothing Then
                            Console.WriteLine("Rtf body:")
                            Console.WriteLine(System.Text.Encoding.UTF8.GetString(items(m).BodyRtf))
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