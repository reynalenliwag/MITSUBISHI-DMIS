VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_PrintStockStat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Stock Status"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PrintStockStat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   3795
   Begin VB.OptionButton opt_accessories 
      Caption         =   "Accessories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   3075
   End
   Begin VB.OptionButton opt_materials 
      Caption         =   "Materials"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   540
      Width           =   3075
   End
   Begin VB.OptionButton opt_parts 
      Caption         =   "Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Value           =   -1  'True
      Width           =   3075
   End
   Begin VB.ComboBox cboDate_Gen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select date from list"
      Top             =   1320
      Width           =   2205
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Negative On Hand"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   420
      TabIndex        =   2
      Top             =   3600
      Width           =   3075
   End
   Begin Crystal.CrystalReport rptPrintStkStat 
      Left            =   450
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Stock Status Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Height          =   795
      Left            =   2730
      MouseIcon       =   "PrintStockStat.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "PrintStockStat.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1770
      Width           =   675
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
      Left            =   2070
      MouseIcon       =   "PrintStockStat.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "PrintStockStat.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1770
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "AS OF:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1380
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_PrintStockStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSTKSTAT                                                         As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim Stock_Option                                   As String
    If Function_Access(LOGID, "Acess_Print", "STOCK STATUS REPORT") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    If IsDate(cboDate_Gen.Text) = True Then
        Set rsSTKSTAT = New ADODB.Recordset
        rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
            Screen.MousePointer = 11
            rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            If opt_accessories.Value = True Then
                Stock_Option = "{stkstat.TYPE}='A'"
            ElseIf opt_materials.Value = True Then
                Stock_Option = "{stkstat.TYPE}='M'"
            Else
                Stock_Option = "{stkstat.TYPE}='P'"
            End If
            
            PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "stock1.rpt", Stock_Option & " AND {stkstat.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1

            Screen.MousePointer = 0
            NEW_LogAudit "V", "STOCK STATUS REPORT", "", "", "", cboDate_Gen, "", ""
        Else
            MsgSpeechBox "Not Yet Generated!"
        End If
    Else
        MsgSpeechBox "Invalid Date Generated!"
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (STOCK STATUS REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "STOCK STATUS REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set rsSTKSTAT = New ADODB.Recordset
    rsSTKSTAT.Open "select date_gen from PMIS_StkStat group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        cboDate_Gen.Clear
        Do While Not rsSTKSTAT.EOF
            cboDate_Gen.AddItem Null2Date(rsSTKSTAT!DATE_GEN)
            rsSTKSTAT.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PrintStockStat = Nothing
    UnloadForm Me
End Sub

