VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSMS_WarRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warranty Report"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_WarRep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   Begin MSComCtl2.DTPicker txtFROM 
      Height          =   315
      Left            =   660
      TabIndex        =   1
      Top             =   780
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52035585
      CurrentDate     =   39665
   End
   Begin Crystal.CrystalReport warrantyrep 
      Left            =   2700
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboWARREP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   4485
   End
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
      Left            =   3930
      MouseIcon       =   "frmCSMS_WarRep.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_WarRep.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   750
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3270
      MouseIcon       =   "frmCSMS_WarRep.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_WarRep.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "View Report Window"
      Top             =   750
      Width           =   675
   End
   Begin MSComCtl2.DTPicker txtTO 
      Height          =   315
      Left            =   660
      TabIndex        =   2
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52035585
      CurrentDate     =   39665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   390
      TabIndex        =   7
      Top             =   1230
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   870
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose a Type Of Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   2370
   End
End
Attribute VB_Name = "frmCSMS_WarRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillCboReport()
    cboWARREP.AddItem "Warranty Listing"
    cboWARREP.AddItem "Approved QIR"
    cboWARREP.AddItem "Waiting for Approval"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim VRANGE                                         As String
    Dim RSTMP                                          As New ADODB.Recordset
    warrantyrep.Reset

    VRANGE = txtFROM.Value & " - " & txtTO.Value
    warrantyrep.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    warrantyrep.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    warrantyrep.Formulas(2) = "printedby = '" & LOGNAME & "'"

    If cboWARREP.Text = "Warranty Listing" Then
        On Error Resume Next
        Screen.MousePointer = 11

        warrantyrep.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        warrantyrep.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        warrantyrep.Formulas(2) = "printedby = '" & LOGNAME & "'"

        If MsgBox("Do you like to Print in Summary Format", vbQuestion + vbYesNo, "CSMS") = vbYes Then
            warrantyrep.WindowTitle = "Warranty Summary Report"
            warrantyrep.ReportTitle = "Warranty Summary Report"
            PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Warrantysumrepor.rpt", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "WARRANTY REPORTS", "", "", "", "WARRANTY LISTING SUMMARY : " & txtFROM & " - " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            warrantyrep.WindowTitle = "Warranty Report"
            warrantyrep.ReportTitle = "Warranty Report Detailed"
            PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Warrantyrepor.rpt", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "WARRANTY REPORTS", "", "", "", "WARRANTY LISTING DETAIL : " & txtFROM & " - " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If

        Screen.MousePointer = 0
    ElseIf cboWARREP.Text = "Approved QIR" Then
        Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_CQIR WHERE (STATUS IS NOT NULL OR STATUS <> 'P')")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            warrantyrep.WindowTitle = "Approved QIR Report"
            warrantyrep.ReportTitle = "Approved QIR Report"
            warrantyrep.Formulas(3) = "VRANGE = '" & VRANGE & "'"
            PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Approved_QIR.rpt", "{CSMS_CQIR.DATEAPPROVED} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {CSMS_CQIR.DATEAPPROVED} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "WARRANTY REPORTS", "", "", "", "APPROVED QIR : " & txtFROM & " - " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            ShowNoRecord
        End If
    Else
        'MsgBox "Report Under Revision", vbInformation, "CSMS"
        Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_CQIR WHERE STATUS = 'P'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            warrantyrep.WindowTitle = "Waiting for Approval"
            warrantyrep.ReportTitle = "Waiting for Approval"
            warrantyrep.Formulas(3) = "VRANGE = '" & VRANGE & "'"

            PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Waiting_for_Approval.rpt", "{CSMS_CQIR.STATUS} = 'P' AND {CSMS_CQIR.TRANDATE} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {CSMS_CQIR.TRANDATE} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "WARRANTY REPORTS", "", "", "", "WAITING FOR APPROVAL : " & txtFROM & " - " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            ShowNoRecord
        End If
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WARRANTY REPORTS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "WARRANTY REPORTS", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillCboReport
    cboWARREP.ListIndex = 0

    txtFROM.Value = firstDay(LOGDATE)
    txtTO.Value = LOGDATE
End Sub

