Attribute VB_Name = "modLoadTree"
'---------------------------------------------------------------------------------------
' Module    :   modLoadTree
' DateTime  :   8/15/2007 08:47
' Purpose   :   Loads the treeview control on frmMain with data from
'               open database
' Parameters:   None
' Returns   :   None
' Side Effects: None
'---------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure :   LoadTree
' DateTime  :   8/15/2007 08:48
' Purpose   :   Loads the treeview control on frmMain with data from
'               open database
' Parameters:   None
' Returns   :   None
' Side Effects: None
' Calls     :   fileDateTime - to get last update
'               getFileSize (modUtils) - to get size of database.
'---------------------------------------------------------------------------------------

Public Sub LoadTree()

On Error GoTo LoadTree_Error
    Dim tblCtr As Long
    Dim fldCtr As Long
    Dim tblTag As String
    Dim nTables As Long
    Dim fldTag As String
    Dim qryCtr As Long
    
    
    Dim tmpRs As Recordset
    
    ' Initialize the tree on frmMain
    Call frmMain.TV1.Nodes.Add(, , "main", "DataBase")
    ' Add a sub item for tables
    Call frmMain.TV1.Nodes.Add("main", tvwChild, "tables", "Tables")
    
    ' Add a sub item for queries
    Call frmMain.TV1.Nodes.Add("main", tvwChild, "queries", "Queries")
    
    ' Add each query to queries sub item
    For qryCtr = 0 To runinfo.appDb.QueryDefs.Count - 1
        Call frmMain.TV1.Nodes.Add("queries", tvwChild, "qry" & Format(qryCtr), runinfo.appDb.QueryDefs(qryCtr).Name)
    Next
    
    ' Initialize tables counter
    nTables = 0
        
    For tblCtr = 0 To runinfo.appDb.TableDefs.Count - 1
        ' Ignore system tables
        If UCase(Mid(runinfo.appDb.TableDefs(tblCtr).Name, 1, 4)) <> "MSYS" Then
            ' Open the recordset
            Set tmpRs = runinfo.appDb.OpenRecordset(runinfo.appDb.TableDefs(tblCtr).Name)
            If tmpRs.RecordCount > 0 Then
                ' Populate the recordset to get recordcount
                tmpRs.MoveLast
                tmpRs.MoveFirst
                ' Save the item to be added
                tblTag = runinfo.appDb.TableDefs(tblCtr).Name & " (" & Format(tmpRs.RecordCount) & " records) [CLICK TO VIEW]"
            Else
                ' Save the item to be added
                tblTag = runinfo.appDb.TableDefs(tblCtr).Name & " (0 records) [CLICK TO VIEW]"
                
            End If
            
            ' Add the table item to the tables tree item
            Call frmMain.TV1.Nodes.Add("tables", tvwChild, "tbl" & Format(tblCtr), tblTag)
            
            ' Increment the tables counter
            nTables = nTables + 1
            
            ' Add the fields for the current table as subitems to
            ' the current table
            For fldCtr = 0 To tmpRs.Fields.Count - 1
                fldTag = Format(fldCtr) & " : " & tmpRs.Fields(fldCtr).Name & getTypeDesc(tmpRs.Fields(fldCtr).Type)
                
                Call frmMain.TV1.Nodes.Add("tbl" & Format(tblCtr), tvwChild, "fld" & Format(tblCtr) & Format(fldCtr), fldTag)
            Next
            ' Close the recordset
            tmpRs.Close
            
            ' Release the recordset object
            Set tmpRs = Nothing
            
        End If
        
        
    Next
    
    ' Update the text boxes on form frmMain
    
    ' Number of tables
    frmMain.txtTables.Text = Format(nTables)
    
    ' Last update
    frmMain.txtLastUpdate.Text = Format(FileDateTime(runinfo.appDbName))
    
    ' File size
    frmMain.txtSize.Text = GetFileSize(runinfo.appDbName)
    
    ' Number of queries
    frmMain.txtQueries.Text = Format(runinfo.appDb.QueryDefs.Count)
    
    ' Expand the tree on frmMain
    frmMain.TV1.Nodes.Item(1).Expanded = True
    


Exit Sub

LoadTree_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(LoadTree) of (modLoadTree.bas).", vbCritical

End Sub
