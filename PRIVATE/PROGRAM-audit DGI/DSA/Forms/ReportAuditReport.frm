VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportAuditReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audit  Report"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "ReportAuditReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1230
      Width           =   2265
   End
   Begin Crystal.CrystalReport rptLogs 
      Left            =   195
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
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
      Left            =   2175
      MouseIcon       =   "ReportAuditReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportAuditReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   1305
      MouseIcon       =   "ReportAuditReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportAuditReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtT 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2205
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   405012481
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtF 
      Height          =   405
      Left            =   1425
      TabIndex        =   3
      Top             =   1665
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      Format          =   405012481
      CurrentDate     =   39203
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Audit Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   270
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   -3180
      Picture         =   "ReportAuditReport.frx":19D0
      Top             =   0
      Width           =   7665
   End
   Begin VB.Label Label3 
      Caption         =   "User Action"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   1260
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "From Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   5
      Top             =   1755
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "To Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   4
      Top             =   2265
      Width           =   885
   End
End
Attribute VB_Name = "frmReportAuditReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
'Function/Feature: Log Report
'Date Started: 07/05/2007 9:45pm
'Last Update:
'Database Updates:
'Who Updated: AXP
'Updating Code:AXP-0707200713:28
'==============================================================

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Function Feature   : DMIS_AUDIT REPORT
'Date               : 7/7/2007
'Last Update        : 7/7/2007
'Database Update    : DMIS_AUDIT DATABASE
'Who Updated        : AXP
'Upating Code       : AXP-0707200713:28
Private Sub cmdPrint_Click()
'Upating Code       : AXP-0707200713:28
    On Error GoTo ErrorCode

    Screen.MousePointer = 11

    Dim FDate                                               As Date
    Dim TDate                                               As Date
    Dim rsLogs                                              As ADODB.Recordset
    Dim RecordSelection                                     As String
    Set rsLogs = New ADODB.Recordset

    FDate = CDate(dtF.Value)
    TDate = CDate(dtT.Value)

    If FDate = TDate Then
        TDate = DateAdd("d", 1, FDate)
    End If

    If UCase(Combo1) = "<ALL>" Then
        If gconAudit.Execute("SELECT COUNT(*) from ALL_vw_Audit where action_date BETWEEN  '" & FDate & "' AND '" & DateAdd("d", 1, TDate) & "'").Fields(0).Value = 0 Then
            ShowNoRecord
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else

        If gconAudit.Execute("SELECT COUNT(*) from ALL_vw_Audit WHERE action_date BETWEEN  '" & FDate & "' AND '" & DateAdd("d", 1, TDate) & "' AND USERACTION='" & Combo1 & "'").Fields(0).Value = 0 Then
            ShowNoRecord
            Screen.MousePointer = 0
            Exit Sub
        End If

    End If


    rptLogs.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptLogs.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    rptLogs.Formulas(2) = "datefrom = '" & FDate & "'"
    rptLogs.Formulas(3) = "dateto = '" & TDate & "'"
    '{A.ACTION_DATE} > DATE(2007,1,2)
    If UCase(Combo1) <> "<ALL>" Then
        RecordSelection = "{A.USERACTION}='" & Combo1 & "' AND ({A.ACTION_DATE}>=DATE(" & Year(dtF) & "," & Month(dtF) & "," & Day(dtF) & ") and {A.ACTION_DATE}<=DATE(" & Year(dtT) & "," & Month(dtT) & "," & Day(dtT) & "))"
    Else
        RecordSelection = "({A.ACTION_DATE}>=DATE(" & Year(dtF) & "," & Month(dtF) & "," & Day(dtF) & ") and {A.ACTION_DATE}<=DATE(" & Year(dtT) & "," & Month(dtT) & "," & Day(dtT) & "))"
    End If
    PrintSQLReport rptLogs, CRIS_REPORT_PATH & "AuditReport.rpt", RecordSelection, DMIS_Audit_Connection, 1
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()

    Screen.MousePointer = 11

    If gconAudit.State = 0 Then
        MsgBox " Cannot Open Audit database Please Configure Server Configuration ", vbCritical
        Screen.MousePointer = 0
        Exit Sub
    End If

    CenterMe frmMain, Me, 0
    Screen.MousePointer = 0
    dtF.Value = firstDay(LOGDATE)
    dtT.Value = LOGDATE
    Dim temprs                                              As ADODB.Recordset
    On Error GoTo ErrorCode:
Loading:     Set temprs = gconAudit.Execute("Select distinct USERACTION FROM ALL_vw_Audit ")

    'set temprs=
    If Not (temprs.EOF Or temprs.BOF) Then
        Combo_Loadval Combo1, temprs

    End If
    Call Combo1.AddItem("<ALL>", 0)
    Combo1.ListIndex = 0
    Exit Sub
ErrorCode:
    On Error GoTo MayError

    If CHANGE_USER = True Then
        gconAudit.Execute "ALTER  VIEW ALL_vw_AUDIT  " & vbCrLf & _
                          " AS " & vbCrLf & _
                          " SELECT A.ACTION_DATE, A.MODULE_NAME AS DESCRIPTION, A.TYPE, A.TRANNO, A.TRANSTYPE, A.DETID, " & _
                          " CASE ISNULL(A.USER_ACTION,'') " & _
                          " WHEN 'A' THEN 'ADDED'" & _
                          " WHEN 'E' THEN 'EDITED'" & _
                          " WHEN 'P' THEN 'POSTED'" & _
                          " WHEN 'U' THEN 'UNPOSTED'" & _
                          " WHEN 'C' THEN 'CANCELLED'" & _
                          " WHEN 'X' THEN 'DELETED'" & _
                          " WHEN 'V' THEN 'VIEWED'" & _
                          " WHEN 'I' THEN 'INQUIRED'" & _
                          " WHEN 'R' THEN 'PROCESSED'" & _
                          " WHEN 'G' THEN 'GENERATED'" & _
                          " WHEN 'O' THEN 'BATCH POSTING'" & _
                          " ELSE '' END AS USERACTION, A.TRACKING_MEMO, A.[ID], C.USER_NAME, A.[USER_ID], A.MODULE_NAME, A.USER_ACTION, A.TRANSACTION_ID" & _
                          " FROM DMIS_AUDIT.DBO.DMIS_AUDIT A INNER JOIN" & _
                          " DMIS.DBO.ALL_RAMS_USERS C ON A.[USER_ID] = C.[USERID]"
    Else
        gconAudit.Execute "ALTER  VIEW ALL_vw_AUDIT  " & vbCrLf & _
                          " AS " & vbCrLf & _
                          " SELECT A.ACTION_DATE, A.MODULE_NAME AS DESCRIPTION, A.TYPE, A.TRANNO, A.TRANSTYPE, A.DETID, " & _
                          " CASE ISNULL(A.USER_ACTION,'') " & _
                          " WHEN 'A' THEN 'ADDED'" & _
                          " WHEN 'E' THEN 'EDITED'" & _
                          " WHEN 'P' THEN 'POSTED'" & _
                          " WHEN 'U' THEN 'UNPOSTED'" & _
                          " WHEN 'C' THEN 'CANCELLED'" & _
                          " WHEN 'X' THEN 'DELETED'" & _
                          " WHEN 'V' THEN 'VIEWED'" & _
                          " WHEN 'I' THEN 'INQUIRED'" & _
                          " WHEN 'R' THEN 'PROCESSED'" & _
                          " WHEN 'G' THEN 'GENERATED'" & _
                          " WHEN 'O' THEN 'BATCH POSTING'" & _
                          " ELSE '' END AS USERACTION, A.TRACKING_MEMO, A.[ID], C.USERNAME, A.[USER_ID], A.MODULE_NAME, A.USER_ACTION, A.TRANSACTION_ID" & _
                          " FROM DMIS_AUDIT.DBO.DMIS_AUDIT A INNER JOIN" & _
                          " DMIS.DBO.ALL_RAMS_USERS C ON A.[USER_ID] = C.[USERID]"
    End If
    GoTo Loading
    Exit Sub

MayError:
    Resume Next
End Sub

