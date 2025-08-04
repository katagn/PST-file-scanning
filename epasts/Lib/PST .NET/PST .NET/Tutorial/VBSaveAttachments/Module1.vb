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

                        For i As Integer = 0 To items(m).Attachments.Count - 1

                            Dim attachment As Attachment = items(m).Attachments(i)

                            Dim fileName = attachment.FileName

                            If fileName Is Nothing Then
                                fileName = attachment.DisplayName
                            End If

                            Dim filePath As String = "c:\temp\" & fileName

                            attachment.Save(filePath, True)
                        Next
                    Next
                End If
            End Using
        End Sub
    End Class
End Namespace