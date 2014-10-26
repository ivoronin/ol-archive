Public Sub Raise(ByVal Error As ErrObject)
    Err.Raise Error.Number, Error.Source, Error.Description, Error.HelpFile, Error.HelpContext
End Sub

Private Function GetOrCreateFolder(ByVal FolderName As String, ByVal ParentFolder As MAPIFolder) As MAPIFolder
    On Error Resume Next
    Set GetOrCreateFolder = ParentFolder.Folders(FolderName)
    If Err.Number = &H8004010F Or Err.Number = 440 Then
        On Error GoTo 0
        ' "Outlook data file cannot be accessed" or "Array out of bounds"
        Set GetOrCreateFolder = ParentFolder.Folders.Add(FolderName)
    ElseIf Err.Number <> 0 Then
        Raise Err
    End If
    On Error GoTo 0
End Function

Sub Archive()
    Dim FolderName As String
    Dim RootFolder As MAPIFolder
    Dim ArchiveFolder As MAPIFolder
    Dim ArchiveYearFolder As MAPIFolder
    Dim ArchiveMonthFolder As MAPIFolder
    
    Set RootFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent
    Set ArchiveFolder = GetOrCreateFolder("Archives", RootFolder)
    Set ArchiveYearFolder = GetOrCreateFolder(Format(Now(), "yyyy"), ArchiveFolder)
    Set ArchiveMonthFolder = GetOrCreateFolder(Format(Now(), "MM"), ArchiveYearFolder)
    
    For Each Msg In ActiveExplorer.Selection
        Msg.Move ArchiveMonthFolder
    Next Msg
End Sub
