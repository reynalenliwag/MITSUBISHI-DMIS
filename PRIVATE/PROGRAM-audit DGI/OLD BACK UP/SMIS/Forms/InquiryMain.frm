VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_InquiryMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INQUIRY"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "InquiryMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin VB.PictureBox picInqList 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7440
      Left            =   0
      ScaleHeight     =   7440
      ScaleWidth      =   1590
      TabIndex        =   1
      Top             =   0
      Width           =   1590
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Sales Executive Performance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   645
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "View Sales Executive Performance"
         Top             =   1770
         Width           =   1425
      End
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Sales Appointment By SAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   645
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "View Sales Appointment by SAE"
         Top             =   1140
         Width           =   1425
      End
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Prospects Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   645
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "View Prospects Inquiry"
         Top             =   510
         Width           =   1425
      End
      Begin VB.Label Label 
         Caption         =   "Inquiry List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   2325
      End
   End
   Begin VB.PictureBox picInquiry 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7440
      Left            =   1590
      ScaleHeight     =   7440
      ScaleWidth      =   10680
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      Begin XtremeReportControl.ReportControl lvInquiry 
         Height          =   2055
         Left            =   180
         TabIndex        =   11
         Top             =   4500
         Width           =   7140
         _Version        =   655364
         _ExtentX        =   12594
         _ExtentY        =   3625
         _StockProps     =   64
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         ShowFooter      =   -1  'True
      End
      Begin VB.PictureBox picInq 
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   1
         Left            =   0
         ScaleHeight     =   870
         ScaleWidth      =   10575
         TabIndex        =   15
         Top             =   1860
         Width           =   10575
         Begin VB.ComboBox cboMSAYear 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   390
            Width           =   2625
         End
         Begin VB.CommandButton cmdPrintLvInq 
            Caption         =   "Print"
            Height          =   375
            Index           =   1
            Left            =   4140
            TabIndex        =   32
            Tag             =   "Sales Appointment Inquiry"
            ToolTipText     =   "Print"
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton cmdPivotVehicles 
            Caption         =   "View"
            Height          =   375
            Left            =   2910
            TabIndex        =   18
            Tag             =   "View Details"
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label lblCap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Sales Appointments By SAE(s) Inquiry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   180
            TabIndex        =   19
            Top             =   60
            Width           =   4965
         End
      End
      Begin VB.PictureBox picInq 
         BorderStyle     =   0  'None
         Height          =   1725
         Index           =   0
         Left            =   30
         ScaleHeight     =   1725
         ScaleWidth      =   10575
         TabIndex        =   3
         Tag             =   "picInq(0)"
         Top             =   0
         Width           =   10575
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   6135
            TabIndex        =   23
            Tag             =   "Status"
            Top             =   480
            Width           =   3135
         End
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   6135
            TabIndex        =   21
            Tag             =   "Classification"
            Top             =   45
            Width           =   3165
         End
         Begin VB.CommandButton cmdPrintLvInq 
            Caption         =   "Print"
            Height          =   375
            Index           =   0
            Left            =   8190
            TabIndex        =   16
            Tag             =   "Sales Appointment Inquiry"
            ToolTipText     =   "Print"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   1380
            TabIndex        =   13
            Tag             =   "LeadSource"
            Top             =   915
            Width           =   3165
         End
         Begin VB.CommandButton cmdInqSAEISales 
            Caption         =   "View"
            Height          =   375
            Left            =   7110
            TabIndex        =   10
            ToolTipText     =   "View Details"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   6135
            TabIndex        =   6
            Tag             =   "MODEL"
            Top             =   840
            Width           =   3150
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   4965
            TabIndex        =   7
            Tag             =   "cboINQMODEL"
            Top             =   855
            Width           =   4515
         End
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   1380
            TabIndex        =   5
            Tag             =   "COLOR"
            Top             =   480
            Width           =   3165
         End
         Begin VB.ComboBox cboInqSAEISales 
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
            Left            =   1380
            TabIndex        =   4
            Tag             =   "SAE"
            Top             =   90
            Width           =   3135
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   720
            TabIndex        =   8
            Tag             =   "cboINQCOLOR"
            Top             =   420
            Width           =   4035
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "SAE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   810
            TabIndex        =   9
            Tag             =   "cboINQSAE"
            Top             =   30
            Width           =   3915
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "LeadSource"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   225
            TabIndex        =   14
            Tag             =   "cboINQLEADSOURCE"
            Top             =   885
            Width           =   4530
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "Classification"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   4980
            TabIndex        =   22
            Tag             =   "cboINQLEADSOURCE"
            Top             =   15
            Width           =   4530
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   4980
            TabIndex        =   24
            Tag             =   "cboINQLEADSOURCE"
            Top             =   450
            Width           =   4530
         End
      End
      Begin VB.PictureBox picInq 
         BorderStyle     =   0  'None
         Height          =   645
         Index           =   2
         Left            =   0
         ScaleHeight     =   645
         ScaleWidth      =   10575
         TabIndex        =   12
         Top             =   2760
         Width           =   10575
         Begin VB.ComboBox cboPerformaceYear 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5490
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   165
            Width           =   1365
         End
         Begin VB.ComboBox cboPerformaceExecutive 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   165
            Width           =   3975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "View"
            Height          =   345
            Left            =   6900
            TabIndex        =   26
            Tag             =   "View Details"
            Top             =   165
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5010
            TabIndex        =   31
            Top             =   150
            Width           =   2115
         End
         Begin VB.Label Label1123 
            BackStyle       =   0  'Transparent
            Caption         =   "SAE Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   30
            Top             =   150
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_InquiryMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GridRs                                                         As ADODB.Recordset
Dim ReportTitle                                                       As String

Sub ResizePics()
    Dim i                                                             As Integer
    For i = 0 To picInq.Count - 1
        If picInq(i).Visible = True Then
            Exit For
        End If
    Next

    lvInquiry.Left = 0
    lvInquiry.Top = picInq(i).Height + 10
    lvInquiry.Height = picInquiry.ScaleHeight - (picInq(i).Top + picInq(i).ScaleHeight)
    lvInquiry.Width = picInquiry.ScaleWidth
End Sub

Sub InitVars()
    ReportControlPaintManager lvInquiry

    Dim i                                                             As Integer
    Dim j                                                             As Integer

    For i = 0 To 5
        cboMSAYear.AddItem Year(LOGDATE) - i
        cboPerformaceYear.AddItem Year(LOGDATE) - i
    Next

    cboMSAYear.ListIndex = 0

    FillCombo "Select DISTINCT SAE from CRIS_PROSPECTS", -1, 0, cboInqSAEISales(0)

    FillCombo "Select DISTINCT COLOR from CRIS_PROSPECTS", -1, 0, cboInqSAEISales(1)

    FillCombo "Select DISTINCT MODEL from CRIS_PROSPECTS", -1, 0, cboInqSAEISales(2)

    FillCombo "Select DISTINCT LeadSource from CRIS_PROSPECTS", -1, 0, cboInqSAEISales(3)

    FillCombo "Select DISTINCT Classification from CRIS_PROSPECTS", -1, 0, cboInqSAEISales(4)


    FillCombo "Select DISTINCT SAE from CRIS_PROSPECTS", -1, 0, cboPerformaceExecutive

    cboPerformaceExecutive.AddItem ("ALL")

    With cboInqSAEISales(5)
        .AddItem "OPEN"
        .AddItem "CLOSED"
        .AddItem "INACTIVE"
        .AddItem "LOST SALES"
    End With


    For i = 0 To picInq.Count - 1
        picInq(i).Move 0, 0, picInquiry.ScaleWidth
    Next

End Sub

Private Sub cboPerformaceExecutive_Click()
    Command1_Click
End Sub

Private Sub cboPerformaceYear_Click()
    Command1.Value = True
End Sub

Private Sub chkInqSAEISales_Click(Index As Integer)

    If chkInqSAEISales(Index).Value = 1 Then
        Call ShadeControl(cboInqSAEISales(Index), True)
        If cboInqSAEISales(Index).ListCount > 0 Then cboInqSAEISales(Index).ListIndex = 0
    Else
        Call ShadeControl(cboInqSAEISales(Index), False)
        cboInqSAEISales(Index).ListIndex = -1

    End If
End Sub

Private Sub Command1_Click()


    On Error GoTo ErrorCode:

    If cboPerformaceExecutive.ListIndex = -1 Or cboPerformaceYear.ListIndex = -1 Then: Exit Sub

    Dim SQL



    If cboPerformaceExecutive.Text = "ALL" Then

        SQL = "SELECT convert(varchar, INVOICEDDATE,101),Salesae,  MODELDESCRIPTION ,CUSTNAME ,COLOR from SMIS_SALESORDER WHERE INVOICEDDATE IS NOT NULL and year(INVOICEDDATE)= @YERR order by 1 desc"
        SQL = Replace(SQL, "@YERR", cboPerformaceYear)
        Call ReportControlAddColumnHeader(lvInquiry, "SN, DateSold,SalesAgent, UnitSold, CustomerName,Color")
        ResizeColumnHeader lvInquiry, "20,40,100,100,100,100"

    Else
        SQL = "SELECT convert(varchar, INVOICEDDATE,101), MODELDESCRIPTION UNITSOLD,CUSTNAME ,COlor from SMIS_SALESORDER WHERE INVOICEDDATE IS NOT NULL and SALESAE ='@SAE' AND year(INVOICEDDATE)= @YERR order by 1 desc"
        SQL = Replace(SQL, "@SAE", cboPerformaceExecutive)
        SQL = Replace(SQL, "@YERR", cboPerformaceYear)
        Call ReportControlAddColumnHeader(lvInquiry, "SN, DateSold, UnitSold, CustomerName,Color")
        ResizeColumnHeader lvInquiry, "20,40,100,100,100 "
    End If

    flex_FillReportView gconDMIS.Execute(SQL), lvInquiry, True


    NEW_LogAudit "V", "PROSPECT INQUIRY", "", "", "", "Sales Executive Performance : SAE NAME:" & cboPerformaceExecutive & ":" & "YEAR" & cboPerformaceExecutive, "", ""
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PROSPECT INQUIRY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PROSPECT INQUIRY", "PRINTING")
            'End If
    End Select

End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub

    If optAdvSearch(0).Value = True Then
        '   DateInquired, Aging, AcctName, Variant, SAE, Color, LeadSource, CLASSIFICATION, Status,ProspectID
        Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(Row.Record(9).Value, Null2String(Row.Record(2).Value))
        frmSMIS_Inquiry_ViewLog.Show
    ElseIf optAdvSearch(1).Value = True Then
        If IsNull(Row.Record(0).Value) = True Then
            Exit Sub
        End If
        Call frmSMIS_Inquiry_ViewLog.SHOWSAEAPPOINTMENTDETAIL(Row.Record(0).Value)
        frmSMIS_Inquiry_ViewLog.Show
    End If

End Sub

Private Sub cmdInqSAEISales_Click()

    Dim TEMPRS                                                        As ADODB.Recordset
    Dim i                                                             As Long
    Dim SearchString1                                                 As String

    For i = 0 To chkInqSAEISales.Count - 1
        If chkInqSAEISales(i).Value = 1 Then
            SearchString1 = SearchString1 & cboInqSAEISales(i).Tag & "='" & cboInqSAEISales(i).Text & "' AND "
        End If
    Next

    If Len(SearchString1) > 0 Then
        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        Set TEMPRS = gconDMIS.Execute("SELECT DateInquired, Aging, AcctName, Variant, SAE, Color, LeadSource, CLASSIFICATION, Status,ProspectID from CRIS_vW_PROSPECT_INQUIRY Where " & SearchString1)
        flex_FillReportView TEMPRS, lvInquiry
    Else

    End If
    NEW_LogAudit "V", "PROSPECT INQUIRY", "", "", "", "PROSPECT INQUIRY :SAE: " & N2Str2Null(cboInqSAEISales(0).Text) & ":" & "COLOR:" & N2Str2Null(cboInqSAEISales(1).Text) & ":" & "LEAD SOURCE:" & N2Str2Null(cboInqSAEISales(2).Text) & ":" & "CLASSIFICATIONS:" & N2Str2Null(cboInqSAEISales(3).Text) & ":" & "STATUS:" & N2Str2Null(cboInqSAEISales(4).Text) & ":" & "MODEL:" & N2Str2Null(cboInqSAEISales(5).Text), "", ""

End Sub

Private Sub cmdPivotVehicles_Click()
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim i                                                             As Long
    Dim SQL                                                           As String
    On Error GoTo ErrorCode:

    If cboMSAYear.Text = vbNullString Then
        MessagePop InfoVoid, "Select Year", "Select Year From The List", 1000, 1
        Exit Sub
    End If



    SQL = " SELECT Sae,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 1 THEN 1 ELSE 0 END) AS January,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 2 THEN 1 ELSE 0 END) AS February,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 3 THEN 1 ELSE 0 END) AS March,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 4 THEN 1 ELSE 0 END) AS April,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 5 THEN 1 ELSE 0 END) AS May,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 6 THEN 1 ELSE 0 END) AS June ,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 7 THEN 1 ELSE 0 END) AS July,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 8 THEN 1 ELSE 0 END) AS August ,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 9 THEN 1 ELSE 0 END) AS September,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 10 THEN 1 ELSE 0 END) AS October,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 11 THEN 1 ELSE 0 END) AS November,  " _
        & " SUM(CASE Month(EndDateTime) WHEN 12 THEN 1 ELSE 0 END) AS December, SAE   " _
        & " From CRIS_SalesAppointments" _
        & " Where Year(EndDateTime)= ' " & cboMSAYear.Text & "'" _
        & " GROUP BY SAE " _
        & " ORDER BY SAE "

    If optAdvSearch(1).Value = True Then
        SQL = Replace(SQL, "@APTYPE", 1)

    Else
        SQL = Replace(SQL, "@APTYPE", 2)

    End If
    Set TEMPRS = gconDMIS.Execute(SQL)



    flex_FillReportView TEMPRS, lvInquiry, False
    Dim ilng(11)                                                      As Long

    For i = 0 To lvInquiry.Records.Count - 1
        ilng(0) = ilng(0) + CLng(lvInquiry.Rows(i).Record(2).Value)
        ilng(1) = ilng(1) + CLng(lvInquiry.Rows(i).Record(3).Value)
        ilng(2) = ilng(2) + CLng(lvInquiry.Rows(i).Record(4).Value)
        ilng(3) = ilng(3) + CLng(lvInquiry.Rows(i).Record(5).Value)
        ilng(4) = ilng(4) + CLng(lvInquiry.Rows(i).Record(6).Value)
        ilng(5) = ilng(5) + CLng(lvInquiry.Rows(i).Record(7).Value)
        ilng(6) = ilng(6) + CLng(lvInquiry.Rows(i).Record(8).Value)
        ilng(7) = ilng(7) + CLng(lvInquiry.Rows(i).Record(9).Value)
        ilng(8) = ilng(8) + CLng(lvInquiry.Rows(i).Record(10).Value)
        ilng(9) = ilng(9) + CLng(lvInquiry.Rows(i).Record(11).Value)
        ilng(10) = ilng(10) + CLng(lvInquiry.Rows(i).Record(12).Value)
        '    ilng(11) = ilng(11) + CLng(lvInquiry.Rows(i).Record(13).Value)

    Next
    lvInquiry.Columns(1).FooterText = "Totals Sales Appointments:"
    lvInquiry.Columns(2).FooterText = ilng(0)
    lvInquiry.Columns(3).FooterText = ilng(1)
    lvInquiry.Columns(4).FooterText = ilng(2)
    lvInquiry.Columns(5).FooterText = ilng(3)
    lvInquiry.Columns(6).FooterText = ilng(4)
    lvInquiry.Columns(7).FooterText = ilng(5)
    lvInquiry.Columns(8).FooterText = ilng(6)
    lvInquiry.Columns(9).FooterText = ilng(7)
    lvInquiry.Columns(10).FooterText = ilng(8)
    lvInquiry.Columns(11).FooterText = ilng(9)
    lvInquiry.Columns(12).FooterText = ilng(10)
    '    lvInquiry.Columns(13).FooterText = ilng(11)

    Erase ilng
    NEW_LogAudit "V", "PROSPECT INQUIRY", "", "", "", "Sales Appointment By SAE INQUIRE BY:" & cboMSAYear, "", ""

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrintLvInq_Click(Index As Integer)
    On Error GoTo ErrorCode:

    If lvInquiry.Records.Count = 0 Then
        MsgSpeechBox "No Record to Print"
        Exit Sub
    End If

    lvInquiry.PrintOptions.Header.Font.Size = "10"
    lvInquiry.PrintOptions.Header.Font.Bold = True
    lvInquiry.PrintOptions.Header.TextCenter = COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & ReportTitle

    lvInquiry.PrintPreview True


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitVars

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

Private Sub optAdvSearch_Click(Index As Integer)
    Dim i                                                             As Integer

    For i = 0 To picInq.Count - 1
        picInq(i).Visible = False
    Next

    If optAdvSearch(Index).Value = True Then
        picInq(Index).Visible = True
        ResizePics
    End If

    lvInquiry.Columns.DeleteAll
    lvInquiry.Records.DeleteAll
    lvInquiry.Populate



    Select Case Index
        Case 0
            ReportTitle = "PROSPECT INQUIRY"
            Call ReportControlAddColumnHeader(lvInquiry, "DATE, AGING, PROSPECTNAME, MODEL, SAE, COLOR, LEADSOURCE, CLASSIFICATION, STATUS")
            Call ResizeColumnHeader(lvInquiry, "50, 35, 100, 100, 80, 80, 80, 80, 80")
            cmdInqSAEISales_Click

        Case 1
            ReportTitle = "MONTHLY SALE APPOINTMENTS BY SAE"
            Call ReportControlAddColumnHeader(lvInquiry, "SAE, JAN, FEB, MARCH, APRIL, MAY, JUN, JUL, AUG,SEP,OCT,NOV,DEC")
            Call ResizeColumnHeader(lvInquiry, "250, 80, 80, 80, 80, 80, 80, 80, 80,80,80,80,80")
            cmdPivotVehicles_Click


            LogAudit "V", "INQUIRY", "Sales Executive Performance -" & " SAE Name: " & cboPerformaceExecutive & " " & cboPerformaceYear    '''RYAN DC CULAWAY MAY 24,2008


        Case 2
            ReportTitle = "Sales Account Executive Performace"
            Call ReportControlAddColumnHeader(lvInquiry, "SN, DateSold, UnitSold, CustomerName,Color")
            If cboPerformaceExecutive.ListCount > 0 Then: cboPerformaceExecutive.ListIndex = 0
            cboPerformaceYear.ListIndex = 0

            LogAudit "V", "INQUIRY", "Sales Appointment By SAE - Monthly Sales Appointments By SAE(s) Inquiry -" & " " & cboMSAYear    '''RYAN DC CULAWAY MAY 24,2008

        Case 3




        Case 4


        Case 5

    End Select
    Me.Caption = "INQUIRY :" & ReportTitle
End Sub

