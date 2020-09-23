Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Procedure : Main
' DateTime  : 8/9/2007 10:13
' Purpose   : Startup procedure
' Parameters: N/A
' Returns   : N/A
' Side Effects: frmMain is displayed
' Calls     : openLast
'---------------------------------------------------------------------------------------

Public Sub Main()
    ' Try to open the last used database
    checkDirs
    
    openLast
    
    ' Display the main form
    frmMain.Show
End Sub
'---------------------------------------------------------------------------------------
' Procedure :   checkDirs
' DateTime  :   8/15/2007 08:49
' Purpose   :   Verifies that necessary directories are available
' Parameters:   None
' Returns   :   N/A
' Side Effects: Creates directories, if necessary.
' Calls     :   N/A
'---------------------------------------------------------------------------------------
Private Sub checkDirs()
    runinfo.baseAppDir = App.Path
    runinfo.baseSettingsDir = App.Path & "\Settings"
    
    
    If Len(Dir(runinfo.baseSettingsDir, vbDirectory)) = 0 Then
        MkDir runinfo.baseSettingsDir
    End If
    
End Sub

