Imports System
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Dim calendar As Folder = file.MailboxRoot.GetFolder("Calendar")

                If calendar IsNot Nothing Then
                    Dim items As IList(Of Item) = calendar.GetItems()

                    For m As Integer = 0 To items.Count - 1
                        If TypeOf items(m) Is Appointment Then
                            Dim appointment As Appointment = DirectCast(items(m), Appointment)

                            Console.WriteLine("Id: " & appointment.Id)
                            Console.WriteLine("Subject: " & appointment.Subject)
                            Console.WriteLine("StartTime: " & appointment.StartTime)
                            Console.WriteLine("EndTime: " & appointment.EndTime)
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