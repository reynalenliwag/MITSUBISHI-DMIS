VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_VehicleMaster 
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
   Icon            =   "Inquiry_VehicleMaster.frx":0000
   LinkTopic       =   "Form1"
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
         Index           =   7
         Left            =   1950
         TabIndex        =   17
         Tag             =   "INVOICEDDATE"
         Top             =   1320
         Width           =   2325
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&Inquiry"
         Height          =   645
         Left            =   11220
         MouseIcon       =   "Inquiry_VehicleMaster.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Inquiry_VehicleMaster.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Search"
         Top             =   120
         Width           =   795
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
         Tag             =   "DateReleased"
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
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "INVOICED DATE"
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
         Index           =   7
         Left            =   540
         TabIndex        =   18
         Tag             =   "cboINQMODEL"
         Top             =   1380
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_VehicleMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GridRs                                                         As ADODB.Recordset
Dim ReportTitle                                                       As String

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
    '  CHECKCOLUMN
    Call ResizeColumnHeader(lvInquiry, "250,40,80,50,40")

    'DESCRIPTIONS, SOURCE,COLOR, C# , V#, E# , INVENTORY STATUS, SALES STATUS, CUSTOMERNAME, ASSIGNEDSAE, AGING
    If Len(SearchString1) > 0 Then
        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        SQL = " SELECT ISNULL(YEER,'') + ISNULL(' ' + MAKE,'') + ISNULL(' ' +  DESCRIPT ,''),  " & _
            " SOURCE , COLOR," & _
            " IGNKEY , VINO, ENGINENO , ISTATUS , " & _
            " SSTATUS, CUSTOMERNAME, ASSIGNEDSAE  FROM SMIS_VW_INQ_VEHICLEMASTER WHERE " & SearchString1
        Set TEMPRS = gconDMIS.Execute(SQL)
        flex_FillReportView TEMPRS, lvInquiry
        
        Call NEW_LogAudit("V", "VEHICLE MASTER INQUIRY", "", "", SearchString1, "", "", "")
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            'If picAdds.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE MASTER INQUIRY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE MASTER INQUIRY", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    ReportControlPaintManager lvInquiry
    FillCombo "SELECT DISTINCT MODEL  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(0)
    FillCombo "SELECT DISTINCT COLOR  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(1)
    FillCombo "SELECT DISTINCT MAKE  FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(2)
    FillCombo "SELECT DISTINCT ASSIGNEDSAE FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(3)
    FillCombo "SELECT DISTINCT DATERELEASED FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(4)
    FillCombo "SELECT DISTINCT ISTATUS FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(5)
    FillCombo "SELECT DISTINCT SSTATUS FROM SMIS_vw_INQ_VEHICLEMASTER", -1, 0, cboInq(6)

    FillCombo "SELECT DISTINCT invoiceddate FROM smis_salesorder", -1, 0, cboInq(7)

    ReportTitle = "VEHICLE MASTER INQUIRY"


    '   SQL = " DESCRIPTION, SOURCE , COLOR, IGNKEY , PRODNO, SERIALNO , VINO, ENGINENO , ISTATUS , SSTATUS, CUSTOMERNAME, ASSIGNEDSAE "


    'Call ReportControlAddColumnHeader(lvInquiry, "MAKE , CLASS, DESCRIPT, MODEL, YEER, SOURCE, COLOR, C#, P#, S#, V#, E#, F#, ISTATUS, SSTATUS, CUSTOMERNAME, ASSIGNEDSAE, AGING")

    Call ReportControlAddColumnHeader(lvInquiry, "DESCRIPTIONS, SOURCE,COLOR, C# , V#, E# , ISTATUS, SSTATUS, CLIENT NAME, ASSIGNED SAE, AGING")
    lvInquiry.GroupsOrder.Add lvInquiry.Columns(1)
    ReportControlPaintManager lvInquiry

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

Private Sub lvInquiry_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Index Mod 2 = 0 And Row.GroupRow = False Then
        Metrics.BackColor = &H8BF1AC
    End If
End Sub

Private Sub lvInquiry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Call frmSMIS_Mis_Filter.ConfigGrid(lvInquiry, 3)
        frmSMIS_Mis_Filter.Show 1
    End If
End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub

End Sub

