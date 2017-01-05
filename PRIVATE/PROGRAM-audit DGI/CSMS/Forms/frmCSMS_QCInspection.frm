VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_QCInspection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quality Control Inspection"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_QCInspection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   11400
   Begin VB.PictureBox picAPP 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   2250
      ScaleHeight     =   4365
      ScaleWidth      =   6795
      TabIndex        =   12
      Top             =   2168
      Visible         =   0   'False
      Width           =   6825
      Begin MSComCtl2.DTPicker dtpEDATE 
         Height          =   315
         Left            =   4590
         TabIndex        =   35
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   39706
      End
      Begin VB.TextBox txtREMARKS 
         Height          =   1005
         Left            =   1290
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "frmCSMS_QCInspection.frx":1082
         Top             =   3270
         Width           =   3885
      End
      Begin VB.CommandButton cmdCloseAP 
         Caption         =   "&Cancel"
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
         Left            =   6000
         MouseIcon       =   "frmCSMS_QCInspection.frx":1088
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_QCInspection.frx":11DA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel Entry"
         Top             =   3420
         Width           =   705
      End
      Begin VB.CommandButton cmdSavePart 
         Caption         =   "&Update"
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
         Left            =   5310
         MouseIcon       =   "frmCSMS_QCInspection.frx":1518
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_QCInspection.frx":166A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Save Entry"
         Top             =   3420
         Width           =   705
      End
      Begin VB.ComboBox cboSTATUS 
         Height          =   315
         ItemData        =   "frmCSMS_QCInspection.frx":19BA
         Left            =   1290
         List            =   "frmCSMS_QCInspection.frx":19CA
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2880
         Width           =   3885
      End
      Begin VB.Label lblINFO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DATE EVALUATED"
         Height          =   195
         Index           =   8
         Left            =   3000
         TabIndex        =   34
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "REMARKS"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   33
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Index           =   6
         Left            =   510
         TabIndex        =   32
         Top             =   2970
         Width           =   690
      End
      Begin VB.Label lblRES 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   5
         Left            =   1290
         TabIndex        =   28
         Top             =   2490
         Width           =   3855
      End
      Begin VB.Label lblRES 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   5460
         TabIndex        =   27
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label lblRES 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   1290
         TabIndex        =   26
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label lblRES 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   1290
         TabIndex        =   25
         Top             =   1110
         Width           =   5355
      End
      Begin VB.Label lblRES 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   24
         Top             =   750
         Width           =   2835
      End
      Begin VB.Label lblRES 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   23
         Top             =   390
         Width           =   1545
      End
      Begin VB.Label labid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   5730
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "TECHNICIAN"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   2580
         Width           =   1110
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "HRS WRK."
         Height          =   195
         Index           =   4
         Left            =   4380
         TabIndex        =   20
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "STD. HRS."
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   19
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "JOB DESC."
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   18
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "JOB CODE"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   17
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         Caption         =   "RO NO."
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   450
         Width           =   630
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   -30
         Width           =   11265
         _Version        =   655364
         _ExtentX        =   19870
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "JOB STATUS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picHD 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   60
      ScaleHeight     =   5745
      ScaleWidth      =   11265
      TabIndex        =   0
      Top             =   30
      Width           =   11295
      Begin XtremeReportControl.ReportControl rptREP 
         Height          =   4665
         Left            =   60
         TabIndex        =   6
         Top             =   750
         Width           =   11145
         _Version        =   655364
         _ExtentX        =   19659
         _ExtentY        =   8229
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.TextBox txtSEARCH 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   11115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DOUBLE CLICK TO DISPLAY JOBS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   5460
         Width           =   2535
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11265
         _Version        =   655364
         _ExtentX        =   19870
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "SEARCH FOR REPAIR ORDER"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picDET 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   60
      ScaleHeight     =   2805
      ScaleWidth      =   11265
      TabIndex        =   1
      Top             =   5850
      Width           =   11295
      Begin MSComctlLib.ListView lsvJOBS 
         Height          =   2145
         Left            =   30
         TabIndex        =   2
         Top             =   360
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "JOB CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "JOB DESCRIPTION"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "LTS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "HRS WRK."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "STATUS"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "TC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "REMARKS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DATE"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblRO 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   9030
         TabIndex        =   31
         Top             =   30
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   165
         Index           =   2
         Left            =   9330
         Shape           =   1  'Square
         Top             =   2550
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "WAITING FOR QC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   9540
         TabIndex        =   11
         Top             =   2550
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "REJECTED JOBS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   3
         Left            =   7860
         TabIndex        =   10
         Top             =   2550
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   165
         Index           =   1
         Left            =   7620
         Shape           =   1  'Square
         Top             =   2550
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PASSED JOBS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   2
         Left            =   5940
         TabIndex        =   9
         Top             =   2550
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   165
         Index           =   0
         Left            =   5730
         Shape           =   1  'Square
         Top             =   2550
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DOUBLE CLICK TO PASSED OR FAILED THE JOBS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   2550
         Width           =   3660
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11265
         _Version        =   655364
         _ExtentX        =   19870
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "JOB DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12582912
         GradientColorDark=   8388608
      End
   End
End
Attribute VB_Name = "frmCSMS_QCInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As ADODB.Recordset

Function TechName(xTECHCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN WHERE LTRIM(RTRIM(TECHNICIAN)) = '" & xTECHCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        TechName = Null2String(RSTMP!TECH_NAME)
    Else
        Set RSTMP = New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT COMPANYNAME FROM CSMS_CONTRACTOR WHERE LTRIM(RTRIM(CODE)) = '" & xTECHCODE & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            TechName = Null2String(RSTMP!CompanyName)
        Else
            TechName = ""
        End If
    End If
    Set RSTMP = Nothing
End Function

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.Field
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

Sub displayRepairOrder()
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptREP, " REPAIR ORDER NO, NAME, PLATE NO., DATE RECORDED, SA NAME, ")
    Call ReportControlPaintManager(rptREP)
    Call ResizeColumnHeader(rptREP, "15, 30, 11, 15, 19, 0")
    Call flex_FillReportView(gconDMIS.Execute("SELECT CSMS_Repor.REP_OR, CSMS_Repor.NIYM, CSMS_Repor.PLATE_NO, CSMS_Repor.DTE_RECD, CSMS_vw_EMPNO.NAYM, " & _
                                            " CSMS_Repor.ID " & _
                                            " FROM CSMS_Repor INNER JOIN " & _
                                            " CSMS_RepairOrder ON CSMS_Repor.REP_OR = CSMS_RepairOrder.RO_No INNER JOIN " & _
                                            " CSMS_vw_EMPNO ON CSMS_Repor.RECD_BY = CSMS_vw_EMPNO.CODE " & _
                                            " WHERE (CSMS_Repor.TRANSTYPE = 'R') AND(CSMS_REPAIRORDER.STATUS = 'Finish Job')"), rptREP)

    Screen.MousePointer = 0
End Sub

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

Sub displayJobs(VREP_OR As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim I                                              As Integer
    Dim cnt                                            As Integer
    Dim vCOLOR                                         As String

    cnt = 1
    lsvJOBS.ListItems.Clear
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & VREP_OR & "' AND LIVIL = '1' AND DONE = 'Y'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvJOBS.ListItems.Add(, , Null2String(RSTMP!DETCDE))
            ITEM.SubItems(1) = Null2String(RSTMP!DETDSC)
            ITEM.SubItems(2) = NumericVal(RSTMP!DET_HRS)
            ITEM.SubItems(3) = NumericVal(RSTMP!HRSWRK)

            If Null2String(RSTMP!Status) = "Q" Or Null2String(RSTMP!Status) = "" Then ITEM.SubItems(4) = "Waiting for QC"
            If Null2String(RSTMP!Status) = "Y" Then ITEM.SubItems(4) = "PASSED"

            'If Null2String(rsTMP!Approve) = "" Then
            '    ITEM.SubItems(4) = "Ready for QC"
            '    vCOLOR = Shape1(2).BackColor
            'Else
            '    ITEM.SubItems(4) = Null2String(rsTMP!Approve)
            '    If Null2String(rsTMP!Approve) = "Rejected" Then
            '        ITEM.SubItems(4) = "Failed in QC"
            '        vCOLOR = Shape1(1).BackColor
            '    Else
            '        ITEM.SubItems(4) = "Passed in QC"
            '        vCOLOR = Shape1(0).BackColor
            '    End If
            'End If

            ITEM.SubItems(5) = Null2String(RSTMP!TechCode)
            ITEM.SubItems(6) = Null2String(RSTMP!ID)
            ITEM.SubItems(7) = Null2String(RSTMP!QC_REMARKS)
            ITEM.SubItems(8) = Null2Date(RSTMP!BACKJOB_SCHED)

            'lsvJOBS.ListItems(cnt).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(1).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(2).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(3).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(4).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(5).ForeColor = vCOLOR
            'lsvJOBS.ListItems(cnt).ListSubItems(6).ForeColor = vCOLOR


            cnt = cnt + 1
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdCloseAP_Click()
    picAPP.Visible = False
    picAPP.ZOrder 1

    picHD.Enabled = True
    picDET.Enabled = True
End Sub

Private Sub cmdSavePart_Click()
    Dim XREMARKS                                       As String
    Dim XDATE                                          As String

    If Function_Access(LOGID, "ACESS_EDIT", "QUALITY INSPECTION") = False Then Exit Sub

    If cboSTATUS.Text = "" Then
        ShowIsRequiredMsg ("Status cannot be blank")
        cboSTATUS.SetFocus
        Exit Sub
    End If

    If cboSTATUS.Text = "NOT CHECKED" Then
        If txtREMARKS.Text = "" Then
            ShowIsRequiredMsg ("Remarks Cannot be Blank why jobs not Checked")
            txtREMARKS.SetFocus
            Exit Sub
        End If
    End If

    XREMARKS = N2Str2Null(txtREMARKS)
    XDATE = N2Str2Null(dtpEDATE.Value)

    If MsgBox("Do you want to Tag this job as " & cboSTATUS & "", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
        'OK - PASSED STATUS
        If cboSTATUS.Text = "OK - PASSED" Then
            SQL_STATEMENT = "UPDATE CSMS_RO_DET SET QC_STATUS = 'Y',DONE  = 'Y',STATUS = 'Y',QC_REMARKS = " & XREMARKS & ",BACKJOB_SCHED = " & XDATE & " WHERE ID = " & labID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        ElseIf cboSTATUS.Text = "NOT CHECKED" Then
            SQL_STATEMENT = "UPDATE CSMS_RO_DET SET QC_STATUS = 'Y',DONE  = 'Y',STATUS = 'Y',QC_REMARKS = " & XREMARKS & ",BACKJOB_SCHED = " & XDATE & " WHERE ID = " & labID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        ElseIf cboSTATUS.Text = "NEED TO EVALUATE" Then
            SQL_STATEMENT = "UPDATE CSMS_RO_DET SET QC_STATUS = 'N',DONE  = 'Y',STATUS = 'Q',QC_REMARKS = " & XREMARKS & ",BACKJOB_SCHED = Null WHERE ID = " & labID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = "UPDATE CSMS_RO_DET SET QC_STATUS = 'N',DONE  = NULL,STATUS = 'J',QC_REMARKS = " & XREMARKS & ",BACKJOB_SCHED = NULL,TECHCODE = NULL,BACKJOB_COUNT = BACKJOB_COUNT + 1 WHERE ID = " & labID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
        End If

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblRES(0)), "REP_OR", "CSMS_REPOR"), "", "JOB CODE : " & lblRES(1) & " - QC", "", labID)
        'NEW LOG AUDIT-----------------------------------------------------

        If CheckAllJobsISDone(N2Str2Null(lblRES(0))) = False Then
            gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = NULL, STATUS = 'Back Job', JStatus = 'B' where RO_No = " & N2Str2Null(lblRES(0)) & ""
        End If

        ShowSuccessFullyUpdated

        Call cmdCloseAP_Click
        Call displayJobs(Null2String(lblRO.Caption))
    End If
End Sub

Function CheckAllJobsISDone(VRO) As Boolean
    Dim RS                                             As ADODB.Recordset
    Set RS = New ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE LIVIL = '1' AND (DONE = 'N' OR DONE ='W' OR DONE IS NULL) and REP_OR = " & VRO & "")
    If RS.EOF And RS.BOF Then
        CheckAllJobsISDone = True
    Else
        CheckAllJobsISDone = False
    End If
    Set RS = Nothing
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            If picAPP.Visible = False Then txtSearch.SetFocus

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    Call displayRepairOrder
End Sub

Private Sub lsvJOBS_DblClick()
    If lsvJOBS.ListItems.Count = 0 Then Exit Sub

    Dim INDEX                                          As Integer
    INDEX = lsvJOBS.SelectedItem.INDEX
    labID.Caption = lsvJOBS.ListItems(INDEX).ListSubItems(6)
    If Null2String(lsvJOBS.ListItems(INDEX).ListSubItems(4)) = "Waiting for QC" Then
        picHD.Enabled = False
        picDET.Enabled = False

        cboSTATUS.ListIndex = 0
        lblRES(0).Caption = lblRO.Caption
        lblRES(1).Caption = lsvJOBS.ListItems(INDEX).Text
        lblRES(2).Caption = lsvJOBS.ListItems(INDEX).ListSubItems(1)
        lblRES(3).Caption = lsvJOBS.ListItems(INDEX).ListSubItems(2)
        lblRES(4).Caption = lsvJOBS.ListItems(INDEX).ListSubItems(3)
        lblRES(5).Caption = TechName(LTrim(RTrim(lsvJOBS.ListItems(INDEX).ListSubItems(5))))

        txtREMARKS.Text = Null2String(LTrim(RTrim(lsvJOBS.ListItems(INDEX).ListSubItems(7))))
        If Null2String(LTrim(RTrim(lsvJOBS.ListItems(INDEX).ListSubItems(8)))) = "" Then
            dtpEDATE.Value = Date
        Else
            dtpEDATE.Value = Null2Date(LTrim(RTrim(lsvJOBS.ListItems(INDEX).ListSubItems(8))))
        End If

        picAPP.Visible = True
        picAPP.ZOrder 0
    ElseIf Null2String(lsvJOBS.ListItems(INDEX).ListSubItems(4)) = "PASSED" Then
        MsgBox "Job Already Passed QC", vbInformation, "CSMS"
        Exit Sub
    Else
        MsgBox "This Job is not yet finish", vbInformation, "CSMS"
        Exit Sub
    End If
End Sub

Private Sub rptREP_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    'On Error Resume Next

    If Row.Record Is Nothing Then: Exit Sub

    lblRO.Caption = Null2String(Row.Record(0).Value)
    Call displayJobs(Null2String(Row.Record(0).Value))
End Sub

Private Sub txtsearch_Change()
    rptREP.FilterText = txtSearch.Text
    rptREP.Populate
End Sub

Private Sub txtSEARCH_GotFocus()
    txtSearch.BackColor = &HC0FFFF
End Sub

Private Sub txtSEARCH_LostFocus()
    txtSearch.BackColor = vbWhite
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

