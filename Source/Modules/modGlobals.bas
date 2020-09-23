Attribute VB_Name = "modGlobals"
'---------------------------------------------------------------------------------------
' Module    : modGlobals
' DateTime  : 7/27/2007 08:09
' Purpose   :   Defines global variables
' Parameters:   N/A
' Returns   :   N/A
' Side Effects: N/A
'---------------------------------------------------------------------------------------
Private Type rInfo
    appDb As Database           ' Current database object
    appDbName As String         ' Name of current database
    tableName As String         ' Name of selected table
    baseAppDir As String        ' Path of application
    baseSettingsDir As String   ' Path of settings
    dbTables As Long            ' Number of tables in database
    dbQueries As Long           ' Number of queries in database
    dbSize As Double            ' Size of database (Mb)
    dbRecs As Long              ' Recordcount of database
End Type
Public runinfo As rInfo         ' System status object

Public Const KB As Long = 1024
Public Const MB As Long = 1024 * KB

