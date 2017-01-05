VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_Reminders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INQUIRY"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Inquiry_Reminders.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin XtremeReportControl.ReportControl lvInquiry 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   1770
      Width           =   12150
      _Version        =   655364
      _ExtentX        =   21431
      _ExtentY        =   9975
      _StockProps     =   64
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      ShowFooter      =   -1  'True
   End
   Begin VB.PictureBox picInq 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   12270
      TabIndex        =   1
      Tag             =   "picInq(0)"
      Top             =   0
      Width           =   12270
      Begin VB.CommandButton cmdView 
         Caption         =   "&Inquiry"
         Height          =   645
         Left            =   11310
         MouseIcon       =   "Inquiry_Reminders.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Inquiry_Reminders.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Search"
         Top             =   150
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "INCLUDE VEHICLE DETAILS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   16
         Top             =   1380
         Width           =   4605
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   7740
         TabIndex        =   15
         Tag             =   "SSTATUS"
         Top             =   1230
         Width           =   3015
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "SALES STATUS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   6300
         TabIndex        =   14
         Tag             =   "cboINQMODEL"
         Top             =   1200
         Width           =   4755
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Tag             =   "MODEL"
         Top             =   120
         Width           =   3075
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Tag             =   "COLOR"
         Top             =   510
         Width           =   3075
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Tag             =   "MAKE"
         Top             =   930
         Width           =   3075
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   7740
         TabIndex        =   4
         Tag             =   "ASSIGNEDSAE"
         Top             =   60
         Width           =   3015
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   7740
         TabIndex        =   3
         Tag             =   "ASSIGNEDSAE"
         Top             =   420
         Width           =   3015
      End
      Begin VB.ComboBox cboInq 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   7740
         TabIndex        =   2
         Tag             =   "ISTATUS"
         Top             =   840
         Width           =   3015
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "MODEL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Tag             =   "cboINQSAE"
         Top             =   180
         Width           =   4065
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "COLOR:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   450
         TabIndex        =   9
         Tag             =   "cboINQCOLOR"
         Top             =   540
         Width           =   4065
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE RELEASED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   6240
         TabIndex        =   12
         Tag             =   "cboINQLEADSOURCE"
         Top             =   420
         Width           =   4785
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "INVENTORY STATUS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   5940
         TabIndex        =   13
         Tag             =   "cboINQLEADSOURCE"
         Top             =   840
         Width           =   5085
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "MAKE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   570
         TabIndex        =   6
         Tag             =   "cboINQMODEL"
         Top             =   990
         Width           =   3945
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "ASSIGNED SAE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   6390
         TabIndex        =   11
         Tag             =   "cboINQLEADSOURCE"
         Top             =   30
         Width           =   4620
      End
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GridRs                                                         As ADODB.Recordset
Dim ReportTitle                                                       As String

Sub CHECKCOLUMN()
    '"MAKE , CLASS, DESCRIPT, MODEL, YEER, SOURCE, COLOR, C#, P#, S#, V#, E#, F#, ISTATUS, SSTATUS, CUSTOMERNAME, ASSIGNEDSAE")
    Dim A                                                             As Boolean
    If Check1.Value = 1 Then
        A = True
    Else
        A = False
    End If
    With lvInquiry
        .Columns(7).Visible = A
        .Columns(8).Visible = A
        .Columns(9).Visible = A
        .Columns(10).Visible = A
        .Columns(11).Visible = A
        .Columns(12).Visible = A
    End With
End Sub

Private Sub Check1_Click()
    CHECKCOLUMN
End Sub

Private Sub chkInq_Click(Index As Integer)
    If chkInq(Index).Value = 1 Then
        Call ShadeControl(cboInq(Index), True)
        If cboInq(Index).ListCount > 0 Then cboInq(Index).ListIndex = 0
    Else
        Call ShadeControl(cboInq(Index), False)
        cboInq(Index).ListIndex = -1

    End If
End Sub

Private Sub CmdView_Click()
    Dim SQL                                                           As String
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim i                                                             As Long
    Dim SearchString1                                                 As String
    On Error GoTo ErrorCode:

    For i = 0 To chkInq.Count - 1
        If chkInq(i).Value = 1 Then
            SearchString1 = SearchString1 & cboInq(i).Tag & "='" & cboInq(i).Text & "' AND "
        End If
    Next
    CHECKCOLUMN
    If Len(SearchString1) > 0 Then
        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        SQL = " SELECT MAKE, CLASS,DESCRIPT , " & _
            " MODEL,YEER , Source , Color," & _
            " IGNKEY , PRODNO, SERIALNO , VINO, " & _
            " ENGINENO , FRAMENO, ISTATUS , " & _
            " SSTATUS, CUSTOMERNAME, ASSIGNEDSAE  FROM SMIS_vw_INQ_VEHICLEMASTER WHERE " & SearchString1
        Set TEMPRS = gconDMIS.Execute(SQL)
        flex_FillReportView TEMPRS, lvInquiry
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrintLvInq_Click(Index As Integer)
    If lvInquiry.Records.Count = 0 Then
        MsgSpeechBox "No Record to Print"
        Exit Sub
    End If
    With lvInquiry
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End With
    lvInquiry.PrintOptions.BlackWhiteContrast = 0
    lvInquiry.PrintOptions.BlackWhitePrinting = True
    lvInquiry.PrintOptions.Header.Font.Size = "14"
    lvInquiry.PrintOptions.Header.TextCenter = ReportTitle

    lvInquiry.PrintPreview True
    With lvInquiry
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
    End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    ReportControlPaintManager lvInquiry
    FillCombo "SELECT DISTINCT MODEL  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(0)
    FillCombo "SELECT DISTINCT COLOR  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(1)
    FillCombo "SELECT DISTINCT MAKE  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(2)
    FillCombo "SELECT DISTINCT ASSIGNEDSAE FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(3)
    FillCombo "SELECT DISTINCT DATERELEASED FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(4)
    FillCombo "SELECT DISTINCT ISTATUS FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(5)
    FillCombo "SELECT DISTINCT SSTATUS FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(6)
    ReportTitle = "VEHICLE MASTER INQUIRY"
    Call ReportControlAddColumnHeader(lvInquiry, "MAKE , CLASS, DESCRIPT, MODEL, YEER, SOURCE, COLOR, C#, P#, S#, V#, E#, F#, ISTATUS, SSTATUS, CUSTOMERNAME, ASSIGNEDSAE")
    CHECKCOLUMN
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lvInquiry.Records.Count > 0 Then
        Call frmSMIS_Mis_Filter.ConfigGrid(lvInquiry, 3)
        frmSMIS_Mis_Filter.Show vbModeless
    ElseIf KeyCode = vbKeyF8 And lvInquiry.Records.Count > 0 Then
        lvInquiry.FilterText = vbNullString
        lvInquiry.Populate
        lvInquiry.Columns(4).FooterText = vbNullString
    End If
End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub

End Sub

