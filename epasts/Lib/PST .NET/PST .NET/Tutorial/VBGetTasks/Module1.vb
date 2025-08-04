Imports System
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Dim tasksFolder As Folder = file.MailboxRoot.GetFolder("Tasks")

                If tasksFolder IsNot Nothing Then

                    Dim items As IList(Of Item) = tasksFolder.GetItems()

                    For m As Integer = 0 To items.Count - 1

                        If TypeOf items(m) Is Task Then
                            Dim task As Task = DirectCast(items(m), Task)

                            Console.WriteLine("Id: " & task.Id)
                            Console.WriteLine("Subject: " & task.Subject)
                            Console.WriteLine("StartDate: " & task.StartDate)
                            Console.WriteLine("DueDate: " & task.DueDate)
                            Console.WriteLine("Owner: " & task.Owner)
                            Console.WriteLine("Body: " & task.Body)
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