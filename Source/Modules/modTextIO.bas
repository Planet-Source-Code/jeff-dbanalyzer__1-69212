Attribute VB_Name = "modTextIO"
'---------------------------------------------------------------------------------------
' Module    : modTextIO
' DateTime  : 8/15/2007 08:40
' Purpose   : Contains subs to control opening last opened database.
' Parameters: N/A
' Returns   : N/A
' Side Effects: If last database found, it is opened, and the user
'               interface is loaded with that database.
'---------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : openLast
' DateTime  : 8/15/2007 08:40
' Purpose   : Finds, and if found, opens last database
' Parameters: N/A
' Returns   : N/A
' Side Effects: Database is opened, and user interface is loaded.
' Calls     : openDb, loadTree
'---------------------------------------------------------------------------------------
Public Sub openLast()
    Dim tFile As String
    Dim tHand As Integer
    Dim dbName As String
    
    tFile = runinfo.baseSettingsDir & "\Last.txt"
    If Len(Dir(tFile)) > 0 Then
        tHand = FreeFile()
        Open tFile For Input As #tHand
        Line Input #tHand, dbName
        If Len(Dir(dbName)) > 0 Then
            openDb dbName
            frmMain.mTools.Enabled = True
            frmMain.TV1.Nodes.Clear
            LoadTree

        End If
        Close #tHand
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : saveLast
' DateTime  : 8/15/2007 08:42
' Purpose   : Saves the name of the open database, upon system exit
' Parameters: None
' Returns   : None
' Side Effects: Creates a file (settingsdir\Last.txt) containing
'               database name.
' Calls     : None
'---------------------------------------------------------------------------------------
Public Sub saveLast()
    Dim tFile As String
    Dim tHand As Integer

    
    tFile = runinfo.baseSettingsDir & "\Last.txt"
    
    
    tHand = FreeFile()
    Open tFile For Output As #tHand
    Print #tHand, runinfo.appDbName
    Close #tHand

End Sub

