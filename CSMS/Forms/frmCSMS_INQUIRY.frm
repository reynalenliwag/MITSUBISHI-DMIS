VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_INQUIRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSMS INQUIRY"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_INQUIRY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   14325
   Begin XtremeReportControl.ReportControl rptCUST 
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   14115
      _Version        =   655364
      _ExtentX        =   24897
      _ExtentY        =   4048
      _StockProps     =   64
   End
   Begin XtremeReportControl.ReportControl rptLIST 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   3870
      Width           =   14115
      _Version        =   655364
      _ExtentX        =   24897
      _ExtentY        =   3413
      _StockProps     =   64
   End
   Begin XtremeReportControl.ReportControl rptDET 
      Height          =   1755
      Left            =   120
      TabIndex        =   10
      Top             =   6690
      Width           =   14085
      _Version        =   655364
      _ExtentX        =   24844
      _ExtentY        =   3096
      _StockProps     =   64
   End
   Begin VB.TextBox txtCUST 
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   210
      Width           =   14115
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   14085
      TabIndex        =   7
      Top             =   5910
      Width           =   14115
      Begin VB.TextBox txtDET 
         Height          =   315
         Left            =   810
         TabIndex        =   13
         Top             =   330
         Width           =   2685
      End
      Begin wizButton.cmd cmd2 
         Height          =   375
         Left            =   13080
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         TX              =   "SEARCH"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_INQUIRY.frx":058A
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   390
         Width           =   570
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14115
         _Version        =   655364
         _ExtentX        =   24897
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "DETAIL OPTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   90
      ScaleHeight     =   735
      ScaleWidth      =   14115
      TabIndex        =   1
      Top             =   2970
      Width           =   14145
      Begin VB.TextBox txtFILTER 
         Height          =   345
         Left            =   8850
         TabIndex        =   16
         Top             =   330
         Width           =   3945
      End
      Begin wizButton.cmd cmd1 
         Height          =   345
         Left            =   13050
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         TX              =   "SEARCH"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_INQUIRY.frx":05A6
      End
      Begin MSComCtl2.DTPicker dtpFRom 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   330
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773441
         CurrentDate     =   39658
      End
      Begin MSComCtl2.DTPicker DTPto 
         Height          =   315
         Left            =   3210
         TabIndex        =   6
         Top             =   330
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773441
         CurrentDate     =   39658
      End
      Begin VB.Label lblCUSCODE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   13500
         TabIndex        =   18
         Top             =   30
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   8100
         TabIndex        =   15
         Top             =   450
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE CREATED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1275
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   14145
         _Version        =   655364
         _ExtentX        =   24950
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "REPAID ORDER OPTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH FOR CUSTOMER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   17
      Top             =   30
      Width           =   1815
   End
End
Attribute VB_Name = "frmCSMS_INQUIRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function FindSACODE(VNAME As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT CODE FROM CSMS_VW_EMPNO WHERE NAYM = '" & VNAME & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindSACODE = RSTMP!code
    End If
    Set RSTMP = Nothing
End Function

Function FindTECHCODE(VNAME As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT TECHNICIAN FROM CSMS_VW_TECHNICIAN WHERE TECH_NAME = '" & VNAME & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindTECHCODE = RSTMP!Technician
    End If
    Set RSTMP = Nothing
End Function

Public Function flex_FillReportView(RS As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As Field
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim I                                              As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Sub displayCustomer()
    Dim SCODE                                          As String
    Screen.MousePointer = 11

    'SCODE = FindSACODE(cboSA)
    txtCUST.Text = ""
    Call ReportControlAddColumnHeader(rptCUST, " CUST. CODE, CUSTOMER NAME, CUSTOMER ADDRESS ")
    Call ReportControlPaintManager(rptCUST)
    Call ResizeColumnHeader(rptCUST, "9, 40, 50")
    Call flex_FillReportView(gconDMIS.Execute("SELECT CUSCDE, ACCTNAME, CUSTOMERADD FROM ALL_CUSTOMER_TABLE ORDER BY CUSCDE"), rptCUST)
    Screen.MousePointer = 0
End Sub

Sub displayRepairOrder(vCUSCODE As String)
    Dim SCODE                                          As String
    Screen.MousePointer = 11

    'SCODE = FindSACODE(cboSA)
    txtFILTER.Text = ""
    Call ReportControlAddColumnHeader(rptLIST, " ,RO NO, FULL NAME, PLATE NO, MODEL, SA NAME, ")
    Call ReportControlPaintManager(rptLIST)
    'rptLIST.GroupsOrder.Add rptLIST.Columns(0)
    Call ResizeColumnHeader(rptLIST, "0, 5, 20, 6, 15, 15, 0")
    Call flex_FillReportView(gconDMIS.Execute("SELECT dbo.CSMS_Repor.DTE_RECD, dbo.CSMS_Repor.REP_OR, dbo.CSMS_Repor.NIYM, dbo.CSMS_Repor.PLATE_NO, dbo.CSMS_Repor.MODEL, dbo.CSMS_vw_EMPNO.NAYM " & _
                                            " FROM dbo.CSMS_Repor INNER JOIN " & _
                                            " dbo.CSMS_vw_EMPNO ON dbo.CSMS_Repor.RECD_BY = dbo.CSMS_vw_EMPNO.CODE WHERE dbo.CSMS_Repor.ACCT_NO = '" & vCUSCODE & "' and dbo.CSMS_Repor.TRANSTYPE = 'R' AND DTE_RECD BETWEEN '" & dtpFRom.Value & "' AND '" & DTPto.Value & "'"), rptLIST)
    Screen.MousePointer = 0
End Sub

Sub DisplayDetails(vREPOR As String)
    Dim TCODE                                          As String
    Dim vLIVIL                                         As String
    Screen.MousePointer = 11

    'TCODE = FindTECHCODE(cboTECH)
    txtDET.Text = ""
    Call ReportControlAddColumnHeader(rptDET, " CODE, DESCRIPTION,  STD. HOURS, HR WORK, AMOUNT, TECH. NAME")
    Call ReportControlPaintManager(rptDET)
    Call ResizeColumnHeader(rptDET, " 10, 15, 5, 5, 10, 15")

    Call flex_FillReportView(gconDMIS.Execute("SELECT dbo.CSMS_Ro_Det.DETCDE, dbo.CSMS_Ro_Det.DETDSC,  dbo.CSMS_Ro_Det.DET_HRS, dbo.CSMS_Ro_Det.HRSWRK, " & _
                                            " dbo.CSMS_Ro_Det.DETAMT , dbo.CSMS_RO_DET.TECHNICIAN " & _
                                            " FROM dbo.CSMS_Ro_Det WHERE DBO.CSMS_RO_DET.REP_OR = '" & vREPOR & "'"), rptDET)

    Screen.MousePointer = 0
End Sub

Private Sub cmd1_Click()
    'displayRepairOrder
End Sub

Private Sub cmd2_Click()
    'DisplayDetails
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    dtpFRom.Value = firstDay(Now)
    DTPto.Value = Date

    displayCustomer
End Sub

Private Sub rptCUST_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    Dim vCUSCDE                                        As String

    If Row.Record Is Nothing Then: Exit Sub

    vCUSCDE = Null2String(Row.Record(0).Value)        'CUSTOMER CODE
    displayRepairOrder (vCUSCDE)
End Sub

Private Sub rptLIST_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    Dim vREPOR                                         As String

    If Row.Record Is Nothing Then: Exit Sub

    vREPOR = Null2String(Row.Record(1).Value)         'REPAIR ORDER NO
    DisplayDetails (vREPOR)
End Sub

Private Sub txtCUST_Change()
    rptCUST.FilterText = txtCUST.Text
    rptCUST.Populate
End Sub

Private Sub txtDET_Change()
    rptDET.FilterText = txtDET.Text
    rptDET.Populate
End Sub

Private Sub txtFILTER_Change()
    rptLIST.FilterText = txtFILTER.Text
    rptLIST.Populate
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

