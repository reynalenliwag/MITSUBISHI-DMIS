VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_InvControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Inventory Control"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportInvControl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   5370
   Begin VB.PictureBox Picture1 
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
      Height          =   885
      Left            =   90
      ScaleHeight     =   885
      ScaleWidth      =   4905
      TabIndex        =   9
      Top             =   1140
      Width           =   4905
      Begin VB.TextBox txtyear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3630
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "9999"
         Top             =   270
         Width           =   975
      End
      Begin VB.ComboBox cbomonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3090
         TabIndex        =   13
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Select Monthly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   210
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   150
      Value           =   -1  'True
      Width           =   4905
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Select Date Ranged"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   210
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   450
      Width           =   4905
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
      Left            =   990
      MouseIcon       =   "ReportInvControl.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvControl.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   2250
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
      Left            =   150
      MouseIcon       =   "ReportInvControl.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvControl.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   2250
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   2700
      Top             =   2490
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
   Begin VB.PictureBox picDateRange 
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   4785
      TabIndex        =   2
      Top             =   1110
      Width           =   4785
      Begin MSComCtl2.DTPicker datepFrom 
         Height          =   345
         Left            =   840
         TabIndex        =   3
         Top             =   330
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57737217
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker datepTo 
         Height          =   345
         Left            =   2940
         TabIndex        =   4
         Top             =   330
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57737217
         CurrentDate     =   38216
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2610
         TabIndex        =   6
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   390
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_InvControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:

    If Len(txtyear.Text) = 4 Or txtyear.Text <> "" Then

        rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptReleased.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"

        Set rsMRRINV = New ADODB.Recordset

        If Opt(0).Value = True Then
             rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereceived) = '" & txtyear.Text & "' and month(datereceived) <= " & What_month(cbomonth), gconDMIS, adOpenForwardOnly, adLockReadOnly

        Else
            rsMRRINV.Open "select * from SMIS_MrrInv WHERE datereceived > = '" & datepFrom & "' and datereceived < = '" & datepTo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

        End If
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            If picDateRange.Visible = True Then
                Screen.MousePointer = 11
                rptReleased.WindowShowGroupTree = False
                rptReleased.WindowTitle = "Vehicle Inventory Report"
            
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "invcontrol.rpt", "((({VEHICLE.datereceived} >= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({VEHICLE.datereceived} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) AND {VEHICLE.Released} = false and {VEHICLE.status} = 'P') ", DMIS_REPORT_Connection, 1
            
                Screen.MousePointer = 0
                
                'UPDATED BY: JUN
                'DATE UPDATED: 0903200 4:26
                'NEW LOG AUDIT----------------------------------------------------------------------------------------
                    Call NEW_LogAudit("V", "MONTHLY INVENTORY CONTROL", "", "", "", "For the Month of: " & cbomonth & " " & "Year: " & txtyear, "", "")
                'NEW LOG AUDIT----------------------------------------------------------------------------------------
            Else
                Screen.MousePointer = 11
                '                    FILTER = "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = false)" & _
                                     '                             " OR " & _
                                     '                             "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = true AND month({VEHICLE.datereleased}) > " & What_month(cboMonth) & ")"
    
                FILTER = "year({VEHICLE.datereceived}) = " & txtyear.Text & " AND month({VEHICLE.datereceived}) = " & What_month(cbomonth) & " AND {VEHICLE.Released} = false and {VEHICLE.status} = 'P' "
    
                rptReleased.WindowShowGroupTree = False
                rptReleased.WindowTitle = "Vehicle Inventory Report"
    
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "invcontrol.rpt", FILTER, DMIS_REPORT_Connection, 1
    
                Screen.MousePointer = 0
                
                'UPDATED BY: JUN
                'DATE UPDATED: 0903200 4:26
                'NEW LOG AUDIT----------------------------------------------------------------------------------------
                    Call NEW_LogAudit("V", "MONTHLY INVENTORY CONTROL", "", "", "", "For the Month of: " & cbomonth & " " & "Year: " & txtyear, "", "")
                'NEW LOG AUDIT----------------------------------------------------------------------------------------
    
                
                'LogAudit "V", "MONTHLY INVENTORY CONTROL", cboMonth & " " & txtYear
            End If
        Else
            MsgSpeechBox "No Record for " & cbomonth.Text & " " & txtyear.Text
            Exit Sub
        End If
    End If

    Exit Sub
ErrorCode:
    MsgBox Err.Description
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MONTHLY INVENTORY CONTROL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MONTHLY INVENTORY CONTROL", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cbomonth
    cbomonth.Text = The_month(Month(LOGDATE))
    txtyear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
    Picture1.Visible = True
    picDateRange.Visible = False
End Sub

Private Sub Opt_Click(Index As Integer)
If Opt(0).Value = True Then
    picDateRange.Visible = False
    Opt(0).BackColor = &HFFFF80
    Opt(1).BackColor = &H8000000F
    Picture1.Visible = True
ElseIf Opt(1).Value = True Then
    picDateRange.Visible = True
    Opt(1).BackColor = &HFFFF80
    Opt(0).BackColor = &H8000000F
    Picture1.Visible = False
    If Opt(1).Value = True Then
       datepFrom = firstDay(LOGDATE)
       datepTo = LOGDATE
    End If
End If
End Sub

