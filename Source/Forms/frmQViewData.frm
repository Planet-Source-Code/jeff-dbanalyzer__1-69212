VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmQViewData 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Data"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUnSplit 
      Caption         =   "UnSplit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6405
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split"
      Height          =   375
      Left            =   4395
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DataControl 
      Height          =   330
      Left            =   1080
      Top             =   8280
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dbViewer 
      Bindings        =   "frmQViewData.frx":0000
      Height          =   7335
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   16761024
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   8520
      Width           =   1455
   End
End
Attribute VB_Name = "frmQViewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()

On Error GoTo cmdClose_Click_Error


    Unload Me
    



Exit Sub

cmdClose_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(cmdClose_Click) of (frmQViewData.frm).", vbCritical

End Sub

Private Sub cmdSplit_Click()

On Error GoTo cmdSplit_Click_Error

    dbViewer.Splits.Add dbViewer.Splits.Count

    Me.cmdUnSplit.Enabled = True
        
    

Exit Sub

cmdSplit_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(cmdSplit_Click) of (frmQViewData.frm).", vbCritical

End Sub

Private Sub cmdUnSplit_Click()
    
On Error GoTo cmdUnSplit_Click_Error
    
    dbViewer.Splits.Remove (1)
    If dbViewer.Splits.Count = 1 Then
        Me.cmdSplit.Enabled = True
        Me.cmdUnSplit.Enabled = False
    End If
    


Exit Sub

cmdUnSplit_Click_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(cmdUnSplit_Click) of (frmQViewData.frm).", vbCritical

End Sub

Private Sub Form_Load()

On Error GoTo Form_Load_Error
    Me.Caption = "View " & runinfo.tableName
    Me.DataControl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & runinfo.appDbName & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    Me.DataControl.RecordSource = "Select * from " & runinfo.tableName & ";"
    Me.DataControl.Refresh

    
    

    


Exit Sub

Form_Load_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Form_Load) of (frmQViewData.frm).", vbCritical

End Sub

