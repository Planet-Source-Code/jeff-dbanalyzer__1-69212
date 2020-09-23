VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "dbAnalyzer"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sBarMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7245
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "8/15/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:01 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtQueries 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Text            =   "0"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtLastUpdate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   6
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtTables 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   1455
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8070
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Query Text:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label labQueryText 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   6495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Queries:"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Last Update:"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Size:"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tables:"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "Tools"
      Begin VB.Menu mFormat 
         Caption         =   "Dump Format"
      End
      Begin VB.Menu mDumpFld 
         Caption         =   "Dump Field Numbers"
      End
      Begin VB.Menu mnuDB 
         Caption         =   "DB Utilities"
         Begin VB.Menu mnuCompress 
            Caption         =   "Compress  DB"
         End
         Begin VB.Menu mnuReset 
            Caption         =   "Reset DB"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

On Error GoTo Form_Load_Error
    ' Disable the tools menu, until a database is open
    Me.mTools.Enabled = False
    
    Me.sBarMain.Panels(1).Text = App.EXEName & " " & Format(App.Major) & "." & Format(App.Minor)
    updateMainBar 4, "Ready"
    

    Exit Sub

Form_Load_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Form_Load) of (frmMain.frm).", vbCritical

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mClose_Click
' DateTime  : 8/9/2007 10:15
' Purpose   : Controls the database closing process
' Parameters: N/A
' Returns   : N/A
' Side Effects: None
' Calls     : closeDb
'---------------------------------------------------------------------------------------

Private Sub mClose_Click()

On Error GoTo mClose_Click_Error
    ' Close the database
    closeDb
    
    ' Cleare the tree
    Me.TV1.Nodes.Clear
    
    ' Disable the tools menu
    Me.mTools.Enabled = False
    
    ' Reset text boxes to null
    Me.txtLastUpdate.Text = ""
    Me.txtQueries.Text = ""
    Me.txtSize.Text = ""
    Me.txtTables.Text = ""
        
                

Exit Sub

mClose_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mClose_Click) of (frmMain.frm).", vbCritical

End Sub

Private Sub mDumpFld_Click()

On Error GoTo mDumpFld_Click_Error
    analyzeDB 1
    
    Shell "notepad.exe " & App.Path & "\dbReport.txt", vbMaximizedFocus
    MsgBox "Format data dumped to " & App.Path & "\dbReport.txt"

Exit Sub

mDumpFld_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mDumpFld_Click) of (frmMain.frm).", vbCritical

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mExit_Click
' DateTime  : 8/9/2007 10:17
' Purpose   : Controls the application exit process
' Parameters: None
' Returns   : N/A
' Side Effects: Application stops
' Calls     : saveLast
'---------------------------------------------------------------------------------------
Private Sub mExit_Click()

On Error GoTo mExit_Click_Error
    saveLast
    
    Unload Me
    End
    

Exit Sub

mExit_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mExit_Click) of (frmMain.frm).", vbCritical

End Sub

Private Sub mFormat_Click()

On Error GoTo mFormat_Click_Error
    analyzeDB 0
    
    Shell "notepad.exe " & App.Path & "\dbReport.txt", vbMaximizedFocus
    MsgBox "Format data dumped to " & App.Path & "\dbReport.txt"
    
    

Exit Sub

mFormat_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mFormat_Click) of (frmMain.frm).", vbCritical

End Sub

Private Sub mnuCompress_Click()
    compactDB False
    
End Sub

Private Sub mnuReset_Click()
    resetDb
End Sub

Private Sub mOpen_Click()

On Error GoTo mOpen_Click_Error
    On Error GoTo badfile
    Me.CommonDialog1.DialogTitle = "Open Database File"
    Me.CommonDialog1.DefaultExt = "*.mdb"
    
    Me.CommonDialog1.InitDir = App.Path
    
    Me.CommonDialog1.Filter = "Project Files (*.mdb) | *.mdb| All Files (*.*) | *.*"
    
    
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        If Len(Dir(Me.CommonDialog1.FileName)) > 0 Then
            openDb Me.CommonDialog1.FileName
            Me.mTools.Enabled = True
            Me.TV1.Nodes.Clear
            LoadTree
        Else
            MsgBox Me.CommonDialog1.FileName & " NOT found!", vbCritical, "File Not Found"
            
        End If
        
        
    

        
        
        
        
    End If
    Exit Sub
badfile:
    MsgBox "Unable to open " & Me.CommonDialog1.FileName
    runinfo.appDbName = ""
    'Resume Next
    Exit Sub
    


Exit Sub

mOpen_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mOpen_Click) of (frmMain.frm).", vbCritical

End Sub

Private Sub TV1_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo TV1_NodeClick_Error
    Dim whereTarg As Long
    Dim iTbl As String
    
    If InStr(1, Node.Key, "qry") > 0 Then
        Me.labQueryText.Caption = runinfo.appDb.QueryDefs(Node.Text).SQL
    Else
        Me.labQueryText.Caption = ""
    End If
    If InStr(1, Node.Key, "tbl") > 0 Then
        iTbl = Node.Text
        whereTarg = InStr(1, iTbl, "(")
        iTbl = Trim(Mid(iTbl, 1, whereTarg - 1))
        runinfo.tableName = iTbl
        frmQViewData.Show vbModal, Me
        
    End If
    
        
        
    

Exit Sub

TV1_NodeClick_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(TV1_NodeClick) of (frmMain.frm).", vbCritical

End Sub
