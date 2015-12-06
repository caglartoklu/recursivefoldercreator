VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' File: Form_FormMain
' Includes all the events, functions and subroutines of the FormMain.


Option Compare Database
Option Explicit


' TODO: 5 test with a folder name with a single quote


' Sub: Form_Load
' Load event for the Form.
Private Sub Form_Load()
End Sub


' Sub: ListRecentlyUsedDirectoryEntries_DblClick
' Double Click event for ListRecentlyUsedDirectoryEntries list box.
' Double clicking an item in the recently used directory entries listbox
' will copy the path to the textbox.
Private Sub ListRecentlyUsedDirectoryEntries_DblClick(Cancel As Integer)
    Dim temp As String
    temp = ListRecentlyUsedDirectoryEntries.ItemData(ListRecentlyUsedDirectoryEntries.ListIndex)
    ' TextFolderPath.Text = temp
    TextFolderPath.Value = temp
End Sub


' Sub: CommandPasteFromClipboard_Click
' Click event for PasteFromClipboard button.
' Pastes the contents of the clipboard to the folder text box.
' Since the textbox is focused first, it does not matter where it has been triggered.
Private Sub CommandPasteFromClipboard_Click()
    TextFolderPath.SetFocus
    DoCmd.RunCommand acCmdPaste
End Sub


' Sub: CommandCopyToClipboard_Click
' Click event for CopyToClipboard button.
' Copies the contents of the folder text box to clipboard.
' Since the textbox is focused first, it does not matter where it has been triggered.
Private Sub CommandCopyToClipboard_Click()
    TextFolderPath.SetFocus
    DoCmd.RunCommand acCmdCopy
End Sub


' Sub: CommandCreateFolder_Click
' Click event for CreateFolder button.
' It simply triggers the folder creation.
Private Sub CommandCreateFolder_Click()
    Call DoCreateFolder
End Sub



' Sub: CommandOpenFolder_Click
' Click event for OpenFolder button.
' It opens the folder with Windows Explorer.
Private Sub CommandOpenFolder_Click()
    Dim folderPath As String
    On Error Resume Next
    folderPath = TextFolderPath.Value
    If FolderExists(folderPath) Then
    Call OpenFolderInExplorer(folderPath)
    Else
        Dim prompt As String
        Dim title As String
        prompt = "Folder not found:" & vbCrlf
        prompt = prompt & folderPath
        title = "Open Folder"
        Call MsgBox(prompt, vbOKOnly, title)
    End If
    On Error GoTo 0
End Sub



' Sub: CommandDeleteSelectedLog_Click
' Click event for DeleteSelectedLog button.
' This will delete the selected log from the database and refreshes the list on the screen.
Private Sub CommandDeleteSelectedLog_Click()
    Dim folderPath As String
    If ListRecentlyUsedDirectoryEntries.ListIndex <> -1 Then
        folderPath = ListRecentlyUsedDirectoryEntries.ItemData(ListRecentlyUsedDirectoryEntries.ListIndex)
        Dim sql As String
        If Len(Strip(folderPath)) > 0 Then
            ' TODO: 5 make it SQL injection proof
            sql = "DELETE FROM DirectoryEntries WHERE Path = '" + folderPath + "'"
            DoCmd.SetWarnings False
            DoCmd.RunSQL sql
            DoCmd.SetWarnings True
            Call FillRecentlyUsedEntries
        End If
    End If
End Sub


' Sub: CommandDeleteAllLogs_Click
' Click event for DeleteAllLogs button.
' This will delete all the logs from the database and refreshes the list on the screen.
Private Sub CommandDeleteAllLogs_Click()
    Dim msg As String
    Dim title As String
    Dim answer As Integer
    title = "Delete All Logs"
    msg = msg & "This will delete all the logs." & vbCrlf
    msg = msg & "Are you sure you want to continue?"
    answer = MsgBox(msg, vbOKCancel, title)
    If answer = vbOK Then
        Dim sql As String
        sql = "DELETE FROM DirectoryEntries"
        DoCmd.SetWarnings False
        DoCmd.RunSQL sql
        DoCmd.SetWarnings True
        Call FillRecentlyUsedEntries
    End If
End Sub


' Sub: DoCreateFolder
' Creates the folder, works as a dispatcher by calling all
' the required subs and functions.
Private Sub DoCreateFolder()
    Dim prompt As String
    Dim title As String
    Dim folderPath As String
    folderPath = TextFolderPath.Value
    ' make sure the folder name is normalized
    folderPath = TameFolderPath(folderPath)
    ' add the folder path to the directory entries log.
    Call AddToDirectoryEntries(folderPath)
    ' remove older directory entries according to DirectoryEntryHistoryCount
    Call DeleteOlderDirectoryEntries
    If FolderExists(folderPath) Then
        prompt = "The folder already exists."
        title = "Create Folder"
        Call MsgBox(prompt, vbOKOnly, title)
    Else
        ' the folder, or just some part of it does not exist, create it.
        Call CreateFolderRecursively(folderPath)
        ' TODO: 5 msgbox the result
    End If
    Call FillRecentlyUsedEntries
End Sub


' Function: TameFolderPath
' Fixes possible problems in a file/folder name, such as trimming the name.
' Since this function is to be used only on Windows, the path separators
' can directly be replaced.
' This function can help with the outputs of some applications which provides
' UNIX style path separators on Windows.
'
' Parameters:
' folderPath - the path of the file/folder name.
'
' Returns:
' The tamed folder name.
Private Function TameFolderPath(ByVal folderPath As String) As String
    folderPath = Strip(folderPath)
    folderPath = Replace(folderPath, "/", "\")
    TameFolderPath = folderPath
End Function


' Function: FolderExists
' Returns True if the folder exists, False otherwise.
' Source: http://allenbrowne.com/func-11.html
'
' Parameters:
' strPath - the path of the folder.
'
' Returns:
' True if the folder exists, False otherwise.
Private Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function


' Sub: AddToDirectoryEntries
' Adds the directory to the directory entry logs.
' folderPath - the path of the folder.
' Called from DoCreateFolder() sub.
Private Sub AddToDirectoryEntries(ByVal folderPath As String)
    ' Adds the current folderPath to DirectoryEntries table.
    Dim query As String
    folderPath = SafeSql(folderPath)
    query = "INSERT INTO DirectoryEntries (Path) VALUES ('" + folderPath + "')"
    DoCmd.SetWarnings False
    DoCmd.RunSQL query
    DoCmd.SetWarnings True
End Sub


' Sub: DeleteOlderDirectoryEntries
' Deletes the oldest directories which are not in the top DirectoryEntryHistoryCount logs.
' Called from DoCreateFolder() sub.
Private Sub DeleteOlderDirectoryEntries()
    ' Deletes the DirectoryEntries table except the last n values.
    Dim query As String
    query = query & "DELETE FROM DirectoryEntries "
    query = query & "WHERE ID NOT IN"
    query = query & "("
    query = query & "SELECT TOP " & DirectoryEntryHistoryCount & " ID "
    query = query & "FROM DirectoryEntries "
    query = query & "ORDER BY ID DESC "
    query = query & ")"
    DoCmd.SetWarnings False
    DoCmd.RunSQL query
    DoCmd.SetWarnings True
End Sub


' Sub: FillRecentlyUsedEntries
' Refreshes the recently used folders list on the screen.
' Called when a directory is deleted from logs or a new one is added.
Public Sub FillRecentlyUsedEntries()
    ListRecentlyUsedDirectoryEntries.Requery
End Sub


' Sub: CleanDatabase
' Cleans the database to be ready for release.
' No function/sub calls this sub, it is used in development.
' The users are provided a "Delete All Logs" button.
Public Sub CleanDatabase()
    ' To call this function from immediate window:
    ' Call Form_FormMain.CleanDatabase()
    Call DeleteFromTable("DirectoryEntries")
End Sub
