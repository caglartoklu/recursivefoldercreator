Attribute VB_Name = "mdlAutoExec"
Option Compare Database
Option Explicit

' The number of items to be kept in folder history.
' The older items are deleted in DeleteOlderDirectoryEntries() which
' is called from DoCreateFolder() everytime a folder is created.
Public Const DirectoryEntryHistoryCount = 100


' Function: Autoexec
' Called from the AutoExec macro
' It links the tables to backend databases.
' Also, it rearranges the Access query QDirectoryEntries according to the contant DirectoryEntryHistoryCount.
' It must be a function, not a sub to be compatible with AutoExec macro.
' https://stackoverflow.com/questions/224041/running-code-before-any-forms-open-in-access
' | mdlAutoExec.DoAutoExec()
Public Function DoAutoExec() As Boolean
    Dim backendAccdbFilePath As String
    backendAccdbFilePath = GetBackendProjectPath("DataBackEnd")
    Call mdlDatabase.ConnectTable("DirectoryEntries", backendAccdbFilePath)
    
    Dim query As String
    query = "SELECT TOP " & DirectoryEntryHistoryCount & " Path FROM DirectoryEntries ORDER BY ID DESC;"
    Call SetQueryOfQueryDef("QDirectoryEntries", query)
    
    DoAutoExec = True
End Function
