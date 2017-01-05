VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMS_DirectoryLisiting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Listing"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_DirectoryLisiting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   3165
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
      Height          =   795
      Left            =   2340
      MouseIcon       =   "frmCSMS_DirectoryLisiting.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_DirectoryLisiting.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cboLETTER 
      Height          =   345
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2985
   End
   Begin Crystal.CrystalReport rptLISTING 
      Left            =   90
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Hyundai Dealer Monthly Performance Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
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
      Height          =   795
      Left            =   1620
      MouseIcon       =   "frmCSMS_DirectoryLisiting.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_DirectoryLisiting.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose a Option"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmCSMS_DirectoryLisiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillCBOLetter()
    cboLETTER.AddItem "ALL"
    cboLETTER.AddItem "CUSTOMER START WITH - A"
    cboLETTER.AddItem "CUSTOMER START WITH - B"
    cboLETTER.AddItem "CUSTOMER START WITH - C"
    cboLETTER.AddItem "CUSTOMER START WITH - D"
    cboLETTER.AddItem "CUSTOMER START WITH - E"
    cboLETTER.AddItem "CUSTOMER START WITH - F"
    cboLETTER.AddItem "CUSTOMER START WITH - G"
    cboLETTER.AddItem "CUSTOMER START WITH - H"
    cboLETTER.AddItem "CUSTOMER START WITH - I"
    cboLETTER.AddItem "CUSTOMER START WITH - J"
    cboLETTER.AddItem "CUSTOMER START WITH - K"
    cboLETTER.AddItem "CUSTOMER START WITH - L"
    cboLETTER.AddItem "CUSTOMER START WITH - M"
    cboLETTER.AddItem "CUSTOMER START WITH - N"
    cboLETTER.AddItem "CUSTOMER START WITH - O"
    cboLETTER.AddItem "CUSTOMER START WITH - P"
    cboLETTER.AddItem "CUSTOMER START WITH - Q"
    cboLETTER.AddItem "CUSTOMER START WITH - R"
    cboLETTER.AddItem "CUSTOMER START WITH - S"
    cboLETTER.AddItem "CUSTOMER START WITH - T"
    cboLETTER.AddItem "CUSTOMER START WITH - U"
    cboLETTER.AddItem "CUSTOMER START WITH - V"
    cboLETTER.AddItem "CUSTOMER START WITH - X"
    cboLETTER.AddItem "CUSTOMER START WITH - Y"
    cboLETTER.AddItem "CUSTOMER START WITH - Z"
    cboLETTER.AddItem "CUSTOMER START WITH - 1"
    cboLETTER.AddItem "CUSTOMER START WITH - 2"
    cboLETTER.AddItem "CUSTOMER START WITH - 3"
    cboLETTER.AddItem "CUSTOMER START WITH - 4"
    cboLETTER.AddItem "CUSTOMER START WITH - 5"
    cboLETTER.AddItem "CUSTOMER START WITH - 6"
    cboLETTER.AddItem "CUSTOMER START WITH - 7"
    cboLETTER.AddItem "CUSTOMER START WITH - 8"
    cboLETTER.AddItem "CUSTOMER START WITH - 9"
    cboLETTER.AddItem "CUSTOMER START WITH - 0"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11

    rptLISTING.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptLISTING.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If cboLETTER.Text = "ALL" Then
        PrintSQLReport rptLISTING, CSMS_REPORT_PATH & "Customer.rpt", "", CSMS_REPORT_CONNECTION, 1
    Else
        PrintSQLReport rptLISTING, CSMS_REPORT_PATH & "Customer.rpt", "LEFT({ALL_CUSMAS.CUSNAM},1) = '" & Right(cboLETTER.Text, 1) & "'", CSMS_REPORT_CONNECTION, 1
    End If

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "CUSTOMER DIRECTORY LISTING", "", "", "", cboLETTER, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Screen.MousePointer = 0
    'LogAudit "V", "CUSTOMER DIRECTORY LISTING - REPORTS "
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER DIRECTORY LISTING)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMER DIRECTORY LISTING", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillCBOLetter
    cboLETTER.Text = "ALL"
End Sub

