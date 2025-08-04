Imports System
Imports System.IO
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Program

        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file
                Dim folders As IList(Of Folder) = file.MailboxRoot.GetFolders(True)

                Dim parents As New Dictionary(Of Long, String)()

                Dim parentFolderPath As String = "c:\temp\eml"
                Dim currentFolderPath As String = parentFolderPath & "\" & file.MailboxRoot.DisplayName

                Directory.CreateDirectory(currentFolderPath)
                parents.Add(file.MailboxRoot.Id, currentFolderPath)

                'Create folder structure
                For i As Integer = 0 To folders.Count - 1
                    Dim currentFolder As Folder = folders(i)

                    parentFolderPath = DirectCast(parents(currentFolder.ParentId), String)
                    currentFolderPath = parentFolderPath & "\" & currentFolder.DisplayName

                    Directory.CreateDirectory(currentFolderPath)
                    parents.Add(currentFolder.Id, currentFolderPath)
                Next

                'Get items and save to folders as .msg files
                For j As Integer = 0 To folders.Count - 1

                    For k As Integer = 0 To folders(j).ChildrenCount - 1 Step 1000
                        Dim items As IList(Of Item) = folders(j).GetItems(k, k + 1000)

                        For m As Integer = 0 To items.Count - 1
                            Dim parentFolder As String = DirectCast(parents(items(m).ParentId), String)
                            Dim fileName As String = GetFileName(items(m).Subject)

                            Dim filePath As String = parentFolder & "\" & fileName

                            If filePath.Length > 244 Then
                                filePath = filePath.Substring(0, 244)
                            End If

                            filePath = filePath & "-" & items(m).Id & ".eml"

                            items(m).GetMessageFile().ConvertToMimeMessage().Save(filePath)
                        Next
                    Next
                Next
            End Using
        End Sub

        Private Shared Function GetFileName(ByVal subject As String) As String

            If subject Is Nothing OrElse subject.Length = 0 Then
                Dim fileName As String = "NoSubject"
                Return fileName
            Else
                Dim fileName As String = ""

                For i As Integer = 0 To subject.Length - 1
                    If subject(i) > Chr(31) AndAlso subject(i) < Chr(127) Then
                        fileName += subject(i)
                    End If
                Next

                fileName = fileName.Replace("\", "_")
                fileName = fileName.Replace("/", "_")
                fileName = fileName.Replace(":", "_")
                fileName = fileName.Replace("*", "_")
                fileName = fileName.Replace("?", "_")
                fileName = fileName.Replace("""", "_")
                fileName = fileName.Replace("<", "_")
                fileName = fileName.Replace(">", "_")
                fileName = fileName.Replace("|", "_")

                Return fileName
            End If
        End Function
    End Class
End Namespace