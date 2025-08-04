Imports System
Imports System.Collections.Generic
Imports Independentsoft.Pst

Namespace Sample
    Class Module1
        Shared Sub Main(ByVal args As String())

            Dim file As New PstFile("c:\testfolder\Outlook.pst")

            Using file

                Dim folders As IList(Of Folder) = file.MailboxRoot.GetFolders()

                For i As Integer = 0 To folders.Count - 1

                    Console.WriteLine("Id: " & folders(i).Id)
                    Console.WriteLine("Name: " & folders(i).DisplayName)
                    Console.WriteLine("Type: " & folders(i).ContainerClass)
                    Console.WriteLine("Item count: " & folders(i).ItemCount)
                    Console.WriteLine("--------------------------------------------------------")

                Next

            End Using

            Console.WriteLine("Press ENTER to exit.")
            Console.Read()

        End Sub
    End Class
End Namespace