Attribute VB_Name = "modDbUtils"
'---------------------------------------------------------------------------------------
' Module    : modDbUtils
' DateTime  : 8/9/2007 10:12
' Purpose   : Contains function related to opening, closing and compressing
'             databases.
'---------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : openDb
' DateTime  : 7/27/2007 08:12
' Purpose   :   Opens the database sent in as a parameter
' Parameters:   Name of database to open
' Returns   :   True if database is opened, false if not
' Side Effects: Resets the main form caption
' Calls     :   closeDb, if a database is already open
'---------------------------------------------------------------------------------------
Public Function openDb(newDbName As String) As Boolean
    If Len(runinfo.appDbName) > 0 Then
        closeDb
    End If
    
    If Len(newDbName) > 0 Then
        
        runinfo.appDbName = newDbName
        Set runinfo.appDb = OpenDatabase(newDbName)
        frmMain.Caption = App.Title & " [" & newDbName & "]"
    Else
        frmMain.Caption = App.Title & " []"
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : closeDb
' DateTime  : 8/9/2007 10:08
' Purpose   : If a database is open, this sub closes it.
' Parameters: None
' Returns   : None
' Side Effects: Resets the main form caption
' Calls     : None
'---------------------------------------------------------------------------------------
Public Sub closeDb()
    
    If Len(runinfo.appDbName) > 0 Then
        runinfo.appDb.Close
        Set runinfo.appDb = Nothing
        
        runinfo.appDbName = ""
    End If
    

    frmMain.Caption = App.Title & "[]"
    
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : resetDb
' DateTime  : 8/9/2007 10:09
' Purpose   : Completely erases all tables in database, and
'             then compresses database
' Parameters: None
' Returns   : None
' Side Effects: Empty, compressed database
' Calls     : compactDB
'---------------------------------------------------------------------------------------
Public Sub resetDb()
    Dim tblCtr As Long
    Dim curTable As String
    Dim tmpName As String

    
    If MsgBox("Reset DB - ARE YOU SURE?", vbYesNo, "Reset Confirm") = vbYes Then

        tmpName = runinfo.appDbName
        
        For tblCtr = 0 To runinfo.appDb.TableDefs.Count - 1
            curTable = runinfo.appDb.TableDefs(tblCtr).Name
            If UCase(Mid(curTable, 1, 4)) <> "MSYS" Then
                runinfo.appDb.Execute "Delete * from [" & curTable & "];"
            End If
            
        Next
        
        compactDB True
        
        Set runinfo.appDb = OpenDatabase(runinfo.appDbName)
        
            
        MsgBox runinfo.appDbName & " reset."
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : compactDB
' DateTime  : 8/9/2007 10:10
' Purpose   : Controls the database compression process
' Parameters: boolean quiet - displays message or not
' Returns   : Nothing
' Side Effects: Compressed database
' Calls     : closeDb, compactor, openDb, loadTree
'---------------------------------------------------------------------------------------
Public Sub compactDB(quiet As Boolean)
    Dim tmpName As String
    Dim tmpName1 As String
    updateMainBar 4, "Compacting " & runinfo.appDbName & "..."
    
    tmpName = runinfo.appDbName
    tmpName1 = runinfo.appDbName
    
    closeDb
    compactor tmpName1
    openDb tmpName
    ' Reload the database
    frmMain.TV1.Nodes.Clear
    LoadTree
    
    If Not quiet Then
        MsgBox runinfo.appDbName & " compressed..", vbOKOnly, "Database Compressed"
    End If
    updateMainBar 4, "Ready"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : compactor
' DateTime  : 8/9/2007 10:11
' Purpose   : Compresses the current database
' Parameters: Database name to compress
' Returns   : Nothing
' Side Effects: None
' Calls     : Nothing
'---------------------------------------------------------------------------------------
Private Sub compactor(strDataBaseName As String)
    On Error GoTo Err_CompactDatabase
    Dim strPath As String
    Dim strPath1 As String
    Dim strPathSize As String
    Dim dbError As Boolean
    'Save Paths for Database
    strPath = strDataBaseName
    strPath1 = Left$(strDataBaseName, Len(strDataBaseName) - 4) & "Backup.mdb"
    'Get Size of File Before Compacting
    
    'Kill the file if it exists
    If Dir(strPath1) <> "" Then
        Kill strPath1
    End If
    
    'Compact Database to New Name
    DBEngine.CompactDatabase strPath, strPath1
    ''Kill the file if it exists
    If Dir(strPath) <> "" Then
        Kill strPath
    End If
    'Compact back to original Name
    DBEngine.CompactDatabase strPath1, strPath
    'Kill the file, no need to save it
    If Dir(strPath1) <> "" Then
        Kill strPath1
    End If

    Exit Sub
    
Err_CompactDatabase:
    
    dbError = True

    
    MsgBox "Error compacting " & strDataBaseName

End Sub

