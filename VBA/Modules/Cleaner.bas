Sub FolderKiller(ByVal folderDirectory As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FolderExists(folderDirectory) Then
        FSO.DeleteFolder folderDirectory, True
    Else
        Exit Sub
    End If
End Sub