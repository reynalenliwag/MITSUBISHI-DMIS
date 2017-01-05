VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_InvDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Inventory "
   ClientHeight    =   1575
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   3255
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportInvDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3255
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1620
      MouseIcon       =   "ReportInvDate.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvDate.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   600
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   750
      MouseIcon       =   "ReportInvDate.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvDate.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   90
      Width           =   1440
   End
   Begin Crystal.CrystalReport rptInvDate 
      Left            =   2520
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Vehicle Inventory"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "frmSMIS_Report_InvDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()

    On Error GoTo ErrorCode
    If IsDate(txtDate.Text) = True Then
        Screen.MousePointer = 11
        gconDMIS.Execute "update SMIS_MrrInv set lastinvdate = '" & txtDate.Text & "' "
        rptInvDate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptInvDate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

        PrintSQLReport rptInvDate, SMIS_REPORT_PATH & "unitinventory.rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "UNIT INVENTORY REPORT", txtDate
        rptInvDate.PageZoom 90
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Invalid Date!"
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtDate.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub txtDate_LostFocus()
    txtDate.Text = Format(txtDate.Text, "Short Date")
End Sub

