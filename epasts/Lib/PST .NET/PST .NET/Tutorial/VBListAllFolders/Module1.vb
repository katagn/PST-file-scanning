Imports System
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file
                Dim folders As IList(Of Folder) = file.MailboxRoot.GetFolders(True)

                Dim parentFolder As String = file.MailboxRoot.DisplayName

                Dim parents As New Dictionary(Of Long, String)()

                parents.Add(file.MailboxRoot.Id, parentFolder)

                For i As Integer = 0 To folders.Count - 1
                    Dim currentFolder As Folder = folders(i)

                    parentFolder = DirectCast(parents(currentFolder.ParentId), String)

                    Dim currentFolderPath As String = parentFolder & "\" & currentFolder.DisplayName
                    parents.Add(currentFolder.Id, currentFolderPath)

                    Console.WriteLine("Id: " & currentFolder.Id)
                    Console.WriteLine("Name: " & currentFolder.DisplayName)
                    Console.WriteLine("Type: " & currentFolder.ContainerClass)
                    Console.WriteLine("Item count: " & currentFolder.ItemCount)
                    Console.WriteLine("Path: " & currentFolderPath)
                    Console.WriteLine("----------------------------------------------------------------")
                Next
            End Using

            Console.WriteLine("Press ENTER to exit.")
            Console.Read()

        End Sub
    End Class
End Namespace