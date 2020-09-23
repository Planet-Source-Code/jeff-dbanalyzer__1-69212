Attribute VB_Name = "modUtils"
'---------------------------------------------------------------------------------------
' Module    : modUtils
' DateTime  : 8/15/2007 08:37
' Purpose   : Contains utility, helper functions and subs
' Parameters: Procedure specific
' Returns   : N/A
' Side Effects: N/A
'---------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : getTypeDesc
' DateTime  : 8/15/2007 08:38
' Purpose   : Returns a string identifying the field type number sent in.
' Parameters: Field type number
' Returns   : String description of field type
' Side Effects: None
' Calls     : None
'---------------------------------------------------------------------------------------
Public Function getTypeDesc(iType As Long) As String

On Error GoTo getTypeDesc_Error
    Dim iRet As String
    
    Select Case iType
        Case Is = dbInteger '3
            iRet = " (Int)"
        Case Is = dbSingle  '6
            iRet = " (Single)"
        Case Is = dbCurrency ' 5
            iRet = " (Currency)"
        Case Is = dbDate  '8
            iRet = " (Date)"
        Case Is = dbBoolean '1
            iRet = " (Bool)"
        Case Is = dbLong ' 4
            iRet = " (Long)"
    
        Case Is = dbDouble  ' 7
        
            iRet = " (Double)"
        Case Is = dbText  ' 10
            iRet = " (String)"
    End Select
    getTypeDesc = iRet
    
    

Exit Function

getTypeDesc_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(getTypeDesc) of (modLoadTree.bas).", vbCritical

End Function
'---------------------------------------------------------------------------------------
' Procedure : GetFileSize
' DateTime  : 8/9/2007 10:06
' Purpose   : Returns the file size of the file sent in a parameter
' Parameters: File name to size
' Returns   : Size of file, in megabytes, as a string
' Side Effects: None
' Calls     : FileLen()
'---------------------------------------------------------------------------------------
Public Function GetFileSize(strFile As String) As String

On Error GoTo GetFileSize_Error
    Dim tmpLen As Variant
    tmpLen = FileLen(strFile)
    GetFileSize = Format(tmpLen / MB, "0.00") & " MB"
    
    Exit Function

GetFileSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetFileSize) of (modAnalyzer.bas).", vbCritical

End Function
Public Sub updateMainBar(ndx As Long, msg As String)
    frmMain.sBarMain.Panels(ndx).Text = msg
    
End Sub

