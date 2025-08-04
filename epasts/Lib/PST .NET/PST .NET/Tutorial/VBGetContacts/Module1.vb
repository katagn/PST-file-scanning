Imports System
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Dim contacts As Folder = file.MailboxRoot.GetFolder("Contacts")

                If contacts IsNot Nothing Then
                    Dim items As IList(Of Item) = contacts.GetItems()

                    For m As Integer = 0 To items.Count - 1
                        If TypeOf items(m) Is Contact Then
                            Dim contact As Contact = DirectCast(items(m), Contact)

                            Console.WriteLine("Id: " & contact.Id)
                            Console.WriteLine("GivenName: " & contact.GivenName)
                            Console.WriteLine("Email1Address: " & contact.Email1Address)
                            Console.WriteLine("Email1DisplayName: " & contact.Email1DisplayName)
                            Console.WriteLine("BusinessPhone: " & contact.BusinessPhone)
                            Console.WriteLine("BusinessAddress: " & contact.BusinessAddress)
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