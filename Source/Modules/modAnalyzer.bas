Attribute VB_Name = "modAnalyzer"
'---------------------------------------------------------------------------------------
' Module    : modAnalyzer
' DateTime  : 8/15/2007 08:34
' Purpose   :   Contains code to create database analysis report
' Parameters:   Sub specific
' Returns   :   N/A
' Side Effects: Creates a report (CSV) in application directory.
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : analyzeDB
' DateTime  : 8/15/2007 08:35
' Purpose   : Creates the analyis CSV file
' Parameters: inType, long, specifying type of report
' Returns   : N/A
' Side Effects: N/A
' Calls     : Nothing
'---------------------------------------------------------------------------------------
Public Sub analyzeDB(inType As Long)

On Error GoTo analyzeDB_Error
    Dim oHand As Integer
    Dim oName As String
    Dim tblCtr As Long
    Dim fldCtr As Long
    Dim curTable As String
    Dim tmpRs As Recordset
    Dim pLine As String
    Dim qryCtr As Long
    
    oName = App.Path & "\dbReport.txt"
    oHand = FreeFile()
    Open oName For Output As #oHand
    
    ' Evaluate tables
    For tblCtr = 0 To runinfo.appDb.TableDefs.Count - 1
        ' Save the current table name
        curTable = runinfo.appDb.TableDefs(tblCtr).Name
        ' if it is not a system table, then evaluate it
        If UCase(Mid(curTable, 1, 4)) <> "MSYS" Then
            Set tmpRs = runinfo.appDb.OpenRecordset(curTable)
            
            For fldCtr = 0 To tmpRs.Fields.Count - 1
                If inType = 1 Then
                    pLine = Format(fldCtr) & "," & curTable & "," & tmpRs.Fields(fldCtr).Name & "," & getTypeDesc(tmpRs.Fields(fldCtr).Type)
                Else
                    pLine = curTable & "," & tmpRs.Fields(fldCtr).Name & "," & Format(tmpRs.Fields(fldCtr).Type)
                End If
                
                Print #oHand, pLine
            Next
            tmpRs.Close
            Set tmpRs = Nothing
            
        End If
    Next
    
    'Evaluate queries
    Print #oHand, " "
    For qryCtr = 0 To runinfo.appDb.QueryDefs.Count - 1
        pLine = "Query " & runinfo.appDb.QueryDefs(qryCtr).Name
        Print #oHand, pLine
        pLine = "SQL = " & runinfo.appDb.QueryDefs(qryCtr).SQL
        Print #oHand, pLine
    Next
    
        
        
    ' Close the output file
    Close #oHand
    
            
    
    
    

    Exit Sub

analyzeDB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(analyzeDB) of (modAnalyzer.bas).", vbCritical

End Sub
