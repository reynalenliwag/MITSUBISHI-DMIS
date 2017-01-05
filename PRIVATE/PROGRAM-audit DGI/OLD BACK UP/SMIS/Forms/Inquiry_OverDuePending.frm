VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_OverDuePending 
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
   Icon            =   "Inquiry_OverDuePending.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin XtremeReportControl.ReportControl lvInquiry 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1890
      Width           =   12240
      _Version        =   655364
      _ExtentX        =   21590
      _ExtentY        =   9763
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
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   12270
      TabIndex        =   1
      Tag             =   "picInq(0)"
      Top             =   0
      Width           =   12270
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   1500
         Width           =   615
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
         TabIndex        =   16
         Tag             =   "PO.PO_NO"
         Top             =   1500
         Width           =   3045
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
         Tag             =   "PO.MODELDESCRIPT"
         Top             =   360
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
         Tag             =   "PO.COLOR"
         Top             =   750
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
         Tag             =   "PO.MODELYEAR"
         Top             =   1140
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
         Tag             =   "PO.SAE"
         Top             =   360
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
         Tag             =   "PO.DATEORDERED"
         Top             =   720
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
         Tag             =   "PO.SOURCE"
         Top             =   1110
         Width           =   3015
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Model:"
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
         Left            =   570
         TabIndex        =   10
         Tag             =   "cboINQSAE"
         Top             =   420
         Width           =   3945
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Color:"
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
         Left            =   600
         TabIndex        =   9
         Tag             =   "cboINQCOLOR"
         Top             =   810
         Width           =   3915
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Ordered:"
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
         Left            =   6450
         TabIndex        =   12
         Tag             =   "cboINQLEADSOURCE"
         Top             =   720
         Width           =   4575
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Source:"
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
         Left            =   6930
         TabIndex        =   13
         Tag             =   "cboINQLEADSOURCE"
         Top             =   1110
         Width           =   4095
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
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
         Left            =   660
         TabIndex        =   6
         Tag             =   "cboINQMODEL"
         Top             =   1230
         Width           =   3855
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Executive:"
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
         Left            =   6270
         TabIndex        =   11
         Tag             =   "cboINQLEADSOURCE"
         Top             =   330
         Width           =   4740
      End
      Begin VB.CheckBox chkInq 
         Alignment       =   1  'Right Justify
         Caption         =   "PO No:"
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
         Index           =   6
         Left            =   7020
         TabIndex        =   17
         Tag             =   "cboINQSAE"
         Top             =   1560
         Width           =   4005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         Height          =   645
         Left            =   11220
         MouseIcon       =   "Inquiry_OverDuePending.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Inquiry_OverDuePending.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print"
         Top             =   1050
         Width           =   795
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&Inquiry"
         Height          =   645
         Left            =   11220
         MouseIcon       =   "Inquiry_OverDuePending.frx":07A3
         MousePointer    =   99  'Custom
         Picture         =   "Inquiry_OverDuePending.frx":08F5
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Search"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Result(s):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   19
         Top             =   1560
         Width           =   1395
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12255
         _Version        =   655364
         _ExtentX        =   21616
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   ":::Pending and Over Due Inquiry:::"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_OverDuePending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GridRs                                          As ADODB.Recordset
Dim ReportTitle                                        As String
Private REPORT_NAME                                    As String
Private SELECTIONQRY                                   As String

Sub ShowOverDueOrders()
    REPORT_NAME = "OVERDUE"

End Sub

Sub ShowPendingOrders()
    REPORT_NAME = "PENDING"

End Sub

Sub ShowServerdOrders()
    REPORT_NAME = "SERVERED"

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
    Dim SQL                                            As String
    Dim TEMPRS                                         As ADODB.Recordset
    Dim i                                              As Long
    Dim SearchString1                                  As String
    On Error GoTo ErrorCode:

    For i = 0 To chkInq.Count - 1
        If chkInq(i).Value = 1 Then
            If cboInq(i) = "<ALL>" Then

            Else
                SearchString1 = SearchString1 & cboInq(i).Tag & "  ='" & cboInq(i).Text & "' AND PO.STATUS = 'P' AND "
            End If

        End If
    Next
    '
    If Len(SearchString1) > 0 Then

        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        SearchString1 = " AND  " & SearchString1
        SQL = "SELECT      "
        SQL = SQL & "PO.PO_NO,  "
        SQL = SQL & "PO.DATEORDERED,  "
        SQL = SQL & "PO.DATERECEIVED,  "
        SQL = SQL & "PO.DATEREQ,  "
        SQL = SQL & "DATEDIFF(DAY, PO.DATEREQ, GETDATE()) AS DAYSELASPED, "
        SQL = SQL & " ( Select CUSTOMERNAME From CRIS_VW_ALLPROFILE where CUSCDE= PO.CUSCDE), "
        SQL = SQL & "PO.MODELDESCRIPT,  "
        SQL = SQL & "PO.MODELYEAR,  "
        SQL = SQL & "PO.SAE,  "
        SQL = SQL & "PO.SOURCE,  "
        SQL = SQL & "PO.COLOR,  "
        SQL = SQL & "PO.ID "
        SQL = SQL & " FROM SMIS_PO PO "
        SQL = SQL & " WHERE " & SELECTIONQRY & SearchString1
        Set TEMPRS = gconDMIS.Execute(SQL)

        flex_FillReportView TEMPRS, lvInquiry
    Else
        SQL = "SELECT      "
        SQL = SQL & "PO.PO_NO,  "
        SQL = SQL & "PO.DATEORDERED,  "
        SQL = SQL & "PO.DATERECEIVED,  "
        SQL = SQL & "PO.DATEREQ,  "
        SQL = SQL & "DATEDIFF(DAY, PO.DATEREQ, GETDATE()) AS DAYSELASPED, "
        SQL = SQL & " ( Select CUSTOMERNAME From CRIS_VW_ALLPROFILE where CUSCDE= PO.CUSCDE), "
        SQL = SQL & "PO.MODELDESCRIPT,  "
        SQL = SQL & "PO.MODELYEAR,  "
        SQL = SQL & "PO.SAE,  "
        SQL = SQL & "PO.SOURCE,  "
        SQL = SQL & "PO.COLOR,  "
        SQL = SQL & "PO.ID "
        SQL = SQL & " FROM SMIS_PO PO "
        SQL = SQL & " WHERE PO.STATUS = 'P' AND" & SELECTIONQRY          '& SearchString1
        Set TEMPRS = gconDMIS.Execute(SQL)
        flex_FillReportView TEMPRS, lvInquiry
    End If

    If REPORT_NAME = "PENDING" Then
        Call NEW_LogAudit("V", "PENDING PO", "", "", SearchString1, "", "", "")
    ElseIf REPORT_NAME = "SERVERED" Then
        Call NEW_LogAudit("V", "SERVED PO", "", "", SearchString1, "", "", "")
    Else
        Call NEW_LogAudit("V", "OverDue PO", "", "", SearchString1, "", "", "")
    End If
    Text1 = lvInquiry.Records.Count


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrorCode:
    FlexGrid_To_Excel lvInquiry, lvInquiry.Rows.Count, lvInquiry.Columns.Count, 8
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Public Sub FlexGrid_To_Excel(TheFlexgrid As ReportControl, TheRows As Integer, TheCols As Integer, Optional GridStyle As Integer = 1, Optional WorkSheetName As String)

    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter

    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If

    On Error Resume Next


    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet

    ' name the worksheet
    With wsXL
        If Not WorkSheetName = "" Then
            .Name = WorkSheetName
        End If
    End With

    wsXL.Cells(1, 1).Value = "PO#"
    wsXL.Cells(1, 2).Value = "ORD.DATE"
    wsXL.Cells(1, 3).Value = "REQ.DATE"
    wsXL.Cells(1, 4).Value = "RR.DATE"
    wsXL.Cells(1, 5).Value = "No.Days"
    wsXL.Cells(1, 6).Value = "CUSTOMER"
    wsXL.Cells(1, 7).Value = "UNIT DESCRIPTION"
    wsXL.Cells(1, 8).Value = "YEAR MODEL"
    wsXL.Cells(1, 9).Value = "SAE"
    wsXL.Cells(1, 10).Value = "SOURCE"
    wsXL.Cells(1, 11).Value = "COLOR"



    Dim i                                              As Integer
    For i = 0 To lvInquiry.Rows.Count
        wsXL.Cells(i + 2, 1).Value = "#" & CStr(lvInquiry.Rows(i).Record(0).Value)
        wsXL.Cells(i + 2, 2).Value = lvInquiry.Rows(i).Record(1).Value
        wsXL.Cells(i + 2, 3).Value = lvInquiry.Rows(i).Record(2).Value
        wsXL.Cells(i + 2, 4).Value = lvInquiry.Rows(i).Record(3).Value
        wsXL.Cells(i + 2, 5).Value = lvInquiry.Rows(i).Record(4).Value
        wsXL.Cells(i + 2, 6).Value = lvInquiry.Rows(i).Record(5).Value
        wsXL.Cells(i + 2, 7).Value = lvInquiry.Rows(i).Record(6).Value
        wsXL.Cells(i + 2, 8).Value = lvInquiry.Rows(i).Record(7).Value
        wsXL.Cells(i + 2, 9).Value = lvInquiry.Rows(i).Record(8).Value
        wsXL.Cells(i + 2, 10).Value = lvInquiry.Rows(i).Record(9).Value
        wsXL.Cells(i + 2, 11).Value = lvInquiry.Rows(i).Record(10).Value
    Next

    Dim RG                                             As Range
    For intCol = 1 To TheCols
        wsXL.Columns(intCol).AutoFit
        Set RG = wsXL.Range("A1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows + 1)
        RG.AutoFormat GridStyle
        With RG.Borders
            .ITEM(xlEdgeTop).Weight = xlThin
            .ITEM(xlEdgeLeft).Weight = xlThin
            .ITEM(xlEdgeRight).Weight = xlThin
            .ITEM(xlEdgeBottom).Weight = xlThin
            .LineStyle = 1
        End With


        'RG.Borders().LineStyle = 1
        'RG.Borders().LineStyle = 1
        'RG.Borders().LineStyle = 1
    Next
    objXL.Visible = True
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            'If picAdds.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If REPORT_NAME = "PENDING" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (PENDING PO)"
                Call frmALL_AuditInquiry.DisplayHistory("", "PENDING PO", "PRINTING")
            ElseIf REPORT_NAME = "SERVERED" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVED PO)"
                Call frmALL_AuditInquiry.DisplayHistory("", "SERVED PO", "PRINTING")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (OverDue PO)"
                Call frmALL_AuditInquiry.DisplayHistory("", "OverDue PO", "PRINTING")
            End If

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    ReportControlPaintManager lvInquiry
    'WHERE DATEDIFF(DAY, DATEREQ, GETDATE()) >=0


    If REPORT_NAME = "OVERDUE" Then
        ReportTitle = "OVER DUE INQUIRY LIST"
        SELECTIONQRY = " DATEDIFF(DAY, DATEREQ, GETDATE()) >=0 AND DateReceived is null  "
    ElseIf REPORT_NAME = "PENDING" Then
        SELECTIONQRY = " ISDATE(DateReceived)=0  "
        ReportTitle = "PENDING INQUIRY LIST"

    ElseIf REPORT_NAME = "SERVERED" Then
        SELECTIONQRY = " ISDATE(DateReceived)=1  "
        ReportTitle = "SERVED PO LIST"
    End If
    FillCombo "SELECT DISTINCT MODELDESCRIPT  FROM SMIS_PO WHERE " & SELECTIONQRY & " AND LEN(MODEL)>0", -1, 0, cboInq(0)
    FillCombo "SELECT DISTINCT COLOR  FROM SMIS_PO WHERE " & SELECTIONQRY & " AND  LEN(COLOR)>0", -1, 0, cboInq(1)
    FillCombo "SELECT DISTINCT MODELYEAR  FROM SMIS_PO WHERE " & SELECTIONQRY & " AND LEN(MODELYEAR)>0", -1, 0, cboInq(2)
    FillCombo "SELECT DISTINCT SAE FROM SMIS_PO WHERE " & SELECTIONQRY & "  AND LEN(SAE)>0", -1, 0, cboInq(3)
    FillCombo "SELECT DISTINCT DateOrdered FROM SMIS_PO WHERE " & SELECTIONQRY, -1, 0, cboInq(4)
    FillCombo "SELECT DISTINCT Source FROM SMIS_PO WHERE " & SELECTIONQRY & " AND  LEN(Source)>0", -1, 0, cboInq(5)
    FillCombo "SELECT DISTINCT PO_NO FROM SMIS_PO WHERE " & SELECTIONQRY, -1, 0, cboInq(6)
    Dim i                                              As Integer
    For i = 0 To 6
        Call cboInq(i).AddItem("<ALL>", 0)
    Next

    Call ReportControlAddColumnHeader(lvInquiry, "PONO , OD, RD, RRD, DAY, CUSTOMER, DESCRIPTION, YR, SAE, SOURCE, COLOR")
    'Ordered, Recieved, Required, Elasped, Customer, Description, Year, SAE, SOURCE, COLOR

    ResizeColumnHeader lvInquiry, "4,5,5,5,3,15,10,3,8,5,10"
    lvInquiry.Columns(4).Alignment = xtpAlignmentCenter
    lvInquiry.Columns(7).Alignment = xtpAlignmentCenter
    lvInquiry.PaintManager.TextFont.Name = "Arial"
    lvInquiry.PaintManager.TextFont.Size = "8"

    ' lvInquiry.Columns(2).Visible = False

    Me.Caption = ReportTitle
    ShortcutCaption1.Caption = "::: " & ReportTitle & " :::"

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

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)

    If Row.Record Is Nothing Then: Exit Sub
    If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_Ordering.Show
    frmSMIS_Trans_Ordering.SearchID (Row.Record(11).Value)
    frmSMIS_Trans_Ordering.ZOrder 0

End Sub

