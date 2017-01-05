VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_AfterSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "After Sales Report -Customer"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_Monthly_AfterSalesCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   4410
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
      Left            =   2685
      MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1140
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
      Left            =   1815
      MouseIcon       =   "Report_Monthly_AfterSalesCustomer.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Report_Monthly_AfterSalesCustomer.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1140
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   4440
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   540
      ScaleHeight     =   1305
      ScaleWidth      =   3825
      TabIndex        =   2
      Top             =   0
      Width           =   3825
      Begin VB.TextBox txtYear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   450
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "9999"
         Top             =   600
         Width           =   2325
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   450
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   90
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Left            =   90
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Left            =   330
         TabIndex        =   5
         Top             =   690
         Width           =   825
      End
   End
   Begin VB.PictureBox picRange 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   -150
      ScaleHeight     =   1365
      ScaleWidth      =   4275
      TabIndex        =   10
      Top             =   60
      Width           =   4275
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   405
         Left            =   1440
         TabIndex        =   11
         Top             =   60
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52428801
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   405
         Left            =   1440
         TabIndex        =   12
         Top             =   540
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   52428801
         CurrentDate     =   38216
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
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
         Left            =   960
         TabIndex        =   14
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
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
         Left            =   690
         TabIndex        =   13
         Top             =   90
         Width           =   675
      End
   End
   Begin VB.PictureBox picYear 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   60
      ScaleHeight     =   1305
      ScaleWidth      =   5145
      TabIndex        =   7
      Top             =   30
      Width           =   5145
      Begin VB.TextBox txtYear1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   555
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "9999"
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   930
         TabIndex        =   9
         Top             =   360
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_AfterSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset

Sub SearchMonth()
    picMonthly.Visible = True
    picRange.Visible = False
    picYear.Visible = False
    Me.Caption = "After Sales Monthly:Yearly Customer Directory By Customer Type"
End Sub

Sub SearchRange()
    picMonthly.Visible = False
    picRange.Visible = True
    picYear.Visible = False
    Me.Caption = "After Sales Report:Ranged Customer Directory By Customer Type"
End Sub

Sub SearchYear()
    picMonthly.Visible = False
    picRange.Visible = False
    picYear.Visible = True
    Me.Caption = "After Sales Report:Yearly Customer Directory By Customer Type"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:

    Set rsMRRINV = New ADODB.Recordset
    rptReleased.WindowShowGroupTree = False
    If picYear.Visible = True Then
        If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
            rsMRRINV.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = '" & txtYear1 & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
                Screen.MousePointer = 11
                'FILTER = "(year({PurchAgree.datereleased}) = " & txtYear.Text & " AND ({MRRINV.Released}) = True) AND ({PurchAgree.STATUS}='P')"
                FILTER = "(year({PurchAgree.datereleased}) = " & txtYear.Text & " AND ({MRRINV.Released}) = True) AND ({PurchAgree.STATUS})='P'"
                rptReleased.WindowTitle = "YEALY CUSTOMERS DIRECTORY- BY CUSTOMER TYPE"
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.Formulas(1) = "ForTheMonth = ' For the " & txtYear & "'"
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "aftersales-customer.rpt", FILTER, DMIS_REPORT_Connection, 1
                'LogAudit "V", rptReleased.WindowTitle & "For the " & txtYear
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "AFTER SALES", "", "", "", "FOR THE YEAR " & txtYear1, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
                Exit Sub
            End If
        End If

    ElseIf picMonthly.Visible = True Then
        If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
            rsMRRINV.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = '" & txtYear.Text & "' and month(datereleased) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
                Screen.MousePointer = 11
                FILTER = "(year({PurchAgree.datereleased}) = " & txtYear.Text & " AND month({PurchAgree.datereleased}) = " & What_month(cboMonth) & " AND ({MRRINV.Released}) = True)"
                rptReleased.WindowTitle = "MONTHLY CUSTOMERS DIRECTORY- BY CUSTOMER TYPE"
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.Formulas(1) = "ForTheMonth = 'For the Month of " & cboMonth & " of " & txtYear & "'"
                LogAudit "V", rptReleased.WindowTitle, "For the Month of " & cboMonth & " of " & txtYear

                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "aftersales-customer.rpt", FILTER, DMIS_REPORT_Connection, 1
                
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "AFTER SALES", "", "", "", cboMonth & " " & txtYear, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
                Exit Sub
            End If
        End If
    ElseIf picRange.Visible = True Then
        'If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        rsMRRINV.Open "select * from SMIS_PurchAgree WHERE datereleased Between " & N2Str2Null(FormatDateTime(dtpFrom.Value, vbShortDate)) & "  and " & N2Str2Null(FormatDateTime(dtpTo.Value, vbShortDate)), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            Screen.MousePointer = 11
            FILTER = "((year({PurchAgree.datereleased}) >=" & Year(dtpFrom) & " AND month({PurchAgree.datereleased}) >= " & Month(dtpFrom) & " AND Day({PurchAgree.datereleased}) >= " & Day(dtpFrom) & ")"
            FILTER = FILTER & " AND " & " (year({PurchAgree.datereleased}) <= " & Year(dtpTo) & " AND month({PurchAgree.datereleased}) <= " & Month(dtpTo) & " AND Day({PurchAgree.datereleased}) <= " & Day(dtpTo) & " )) "

            'AND ({MRRINV.Released}) = True)"
            rptReleased.WindowTitle = "RANGED -CUSTOMERS DIRECTORY- BY CUSTOMER TYPE"
            rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptReleased.Formulas(1) = "ForTheMonth = ' For The Period from :" & FormatDateTime(dtpFrom, vbShortDate) & " To " & FormatDateTime(dtpTo, vbShortDate) & "'"

            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "aftersales-customer.rpt", FILTER, DMIS_REPORT_Connection, 1
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
             Call NEW_LogAudit("V", "AFTER SALES", "", "", "", "FROM " & dtpFrom & " " & "TO " & dtpTo, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            'LogAudit "V", rptReleased.WindowTitle, "For The Period from :" & FormatDateTime(dtpFrom, vbShortDate) & " To " & FormatDateTime(dtpTo, vbShortDate)
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
            Exit Sub
        End If
    End If
    '   End If
    
    

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (AFTER SALES)"
            Call frmALL_AuditInquiry.DisplayHistory("", "AFTER SALES", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    txtYear1.Text = Year(LOGDATE)
    dtpFrom.Value = firstDay(LOGDATE)
    dtpTo.Value = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Option1_Click()

End Sub

