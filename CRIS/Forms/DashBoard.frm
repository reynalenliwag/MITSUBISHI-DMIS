VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_Inquiry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "::::INQUIRY::::"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAdvSearch 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   -45
      ScaleHeight     =   8775
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   0
      Width           =   2595
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Loan Application Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   1710
         Width           =   2430
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
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   0
         Left            =   180
         TabIndex        =   29
         Top             =   630
         Value           =   -1  'True
         Width           =   2430
      End
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Sales Order Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   1395
         Width           =   2430
      End
      Begin VB.OptionButton optAdvSearch 
         Caption         =   "Sales Appointments Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   990
         Width           =   2430
      End
      Begin XtremeShortcutBar.ShortcutCaption capInquiry 
         Height          =   450
         Left            =   90
         TabIndex        =   24
         Top             =   0
         Width           =   3465
         _Version        =   655364
         _ExtentX        =   6112
         _ExtentY        =   794
         _StockProps     =   14
         Caption         =   "Inquiry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox picInquiry 
      Align           =   4  'Align Right
      Height          =   8835
      Left            =   2910
      ScaleHeight     =   8775
      ScaleWidth      =   10500
      TabIndex        =   0
      Top             =   0
      Width           =   10560
      Begin XtremeReportControl.ReportControl lvInquiry 
         Height          =   6765
         Left            =   0
         TabIndex        =   12
         Top             =   2070
         Width           =   10410
         _Version        =   655364
         _ExtentX        =   18362
         _ExtentY        =   11933
         _StockProps     =   64
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         ShowFooter      =   -1  'True
      End
      Begin VB.PictureBox picInq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1890
         Index           =   0
         Left            =   45
         ScaleHeight     =   1860
         ScaleWidth      =   10545
         TabIndex        =   3
         Tag             =   "picInq(0)"
         Top             =   45
         Width           =   10575
         Begin VB.CommandButton cmdPrintLvInq 
            Caption         =   "Print"
            Height          =   375
            Index           =   0
            Left            =   6120
            TabIndex        =   22
            Tag             =   "Sales Appointment Inquiry"
            Top             =   1305
            Width           =   1095
         End
         Begin VB.ComboBox cboInqSAEISales 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   345
            Index           =   3
            Left            =   1380
            TabIndex        =   15
            Tag             =   "LeadSource"
            Top             =   1290
            Width           =   3165
         End
         Begin VB.CommandButton cmdInqSAEISales 
            Caption         =   "View"
            Height          =   375
            Left            =   4950
            TabIndex        =   11
            Top             =   1305
            Width           =   1095
         End
         Begin VB.ComboBox cboInqSAEISales 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   345
            Index           =   2
            Left            =   1380
            TabIndex        =   7
            Tag             =   "MODEL"
            Top             =   900
            Width           =   3165
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H8000000D&
            Height          =   405
            Index           =   2
            Left            =   690
            TabIndex        =   8
            Tag             =   "cboINQMODEL"
            Top             =   870
            Width           =   4065
         End
         Begin VB.ComboBox cboInqSAEISales 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   345
            Index           =   1
            Left            =   1380
            TabIndex        =   6
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   345
            Index           =   0
            Left            =   1380
            TabIndex        =   4
            Tag             =   "SAE"
            Top             =   60
            Width           =   3135
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H8000000D&
            Height          =   405
            Index           =   1
            Left            =   720
            TabIndex        =   9
            Tag             =   "cboINQCOLOR"
            Top             =   420
            Width           =   4035
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H8000000D&
            Height          =   405
            Index           =   0
            Left            =   810
            TabIndex        =   10
            Tag             =   "cboINQSAE"
            Top             =   30
            Width           =   3915
         End
         Begin VB.CheckBox chkInqSAEISales 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H8000000D&
            Height          =   405
            Index           =   3
            Left            =   225
            TabIndex        =   16
            Tag             =   "cboINQLEADSOURCE"
            Top             =   1260
            Width           =   4530
         End
      End
      Begin VB.PictureBox picInq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2025
         Index           =   1
         Left            =   -45
         ScaleHeight     =   1995
         ScaleWidth      =   10545
         TabIndex        =   13
         Top             =   45
         Width           =   10575
         Begin VB.CommandButton cmdPrintLvInq 
            Caption         =   "Print"
            Height          =   375
            Index           =   1
            Left            =   9210
            TabIndex        =   23
            Tag             =   "Sales Executive Schedules"
            Top             =   60
            Width           =   1095
         End
         Begin VB.ComboBox cboINQYear1 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   330
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   2625
         End
         Begin VB.CommandButton Command1 
            Caption         =   "View"
            Height          =   375
            Left            =   9240
            TabIndex        =   14
            Top             =   510
            Width           =   1095
         End
         Begin MSComctlLib.ListView lvINQList1 
            Height          =   1545
            Left            =   900
            TabIndex        =   18
            Top             =   360
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CheckBox chkInqSAEIMonths 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Months"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   300
            Width           =   3525
         End
         Begin VB.CheckBox chkInqSAEIYear 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Index           =   0
            Left            =   225
            TabIndex        =   20
            Top             =   0
            Width           =   3525
         End
      End
      Begin VB.PictureBox picInq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2025
         Index           =   2
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   10545
         TabIndex        =   21
         Top             =   0
         Width           =   10575
         Begin VB.CommandButton cmdPivotVehicles 
            Caption         =   "View"
            Height          =   375
            Left            =   3015
            TabIndex        =   26
            Top             =   585
            Width           =   1875
         End
         Begin VB.ComboBox cboMSAYear 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   630
            Width           =   2625
         End
         Begin SHDocVwCtl.WebBrowser wGraphs 
            Height          =   1965
            Left            =   5430
            TabIndex        =   28
            Top             =   0
            Width           =   5115
            ExtentX         =   9022
            ExtentY         =   3466
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin VB.Label lblCap 
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
            ForeColor       =   &H00126322&
            Height          =   915
            Left            =   210
            TabIndex        =   27
            Top             =   60
            Width           =   3165
         End
      End
   End
End
Attribute VB_Name = "frmCRIS_Inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GridRs                                As ADODB.Recordset
Dim ReportTitle As String


Private Sub chkInqSAEIMonths_Click(Index As Integer)
'    If (chkInqSAEIMonths(Index).Value = 1) Then
'        lvlist(Index).Enabled = True
'        lvlist(Index).BackColor = vbWhite
'    Else
'        lvlist(Index).Enabled = False
'        lvlist(Index).BackColor = vbButtonFace
'    End If
End Sub




Private Sub chkProspectInq_Click(Index As Integer)
    If chkProspectInq(Index).Value = 1 And cboFilterInq(Index).ListCount > 0 Then
        cboFilterInq(Index).Enabled = True
        cboFilterInq(Index).BackColor = vbWhite
        cboFilterInq(Index).ListIndex = 0
    ElseIf chkProspectInq(Index).Value = 1 And cboFilterInq(Index).ListCount = 0 Then
        chkProspectInq(Index).Value = 0
    Else
        cboFilterInq(Index).Enabled = False
        cboFilterInq(Index).BackColor = vbButtonFace
    End If
End Sub

Private Sub chkInqSAEISales_Click(Index As Integer)

    
    
    If chkInqSAEISales(Index).Value = 1 Then
        cboInqSAEISales(Index).Enabled = True
        cboInqSAEISales(Index).BackColor = vbWhite
    Else
    
        cboInqSAEISales(Index).Enabled = False
        cboInqSAEISales(Index).BackColor = vbButtonFace
    
        
    End If
End Sub

Private Sub chkInqSAEISalesMonth_Click()
If chkInqSAEISalesMonth.Value = 1 Then
    lvINQList.Enabled = True
Else
    lvINQList.Enabled = False
End If
End Sub

Private Sub cmdInqSAEISales_Click()

    Dim temprs                               As ADODB.Recordset
    Dim FLD                                  As ADODB.Field
    Dim i                                    As Long
    'Dim SearchString2                        As String
    Dim SearchString1                        As String

    For i = 0 To chkInqSAEISales.Count - 1
        If chkInqSAEISales(i).Value = 1 Then
            SearchString1 = SearchString1 & cboInqSAEISales(i).Tag & "='" & cboInqSAEISales(i).Text & "' AND "
        End If
    Next
    If Len(SearchString1) > 0 Then
        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        Set temprs = gconDMIS.Execute("Select ID, AcctName,  DateInquired, AcctName , LeadSource ,VehicleCode , MODEL, Color, SAE  from  CRIS_INQUIRY  Where " & SearchString1)
    Else
        Set temprs = gconDMIS.Execute("Select ID, AcctName,  DateInquired, AcctName , LeadSource ,VehicleCode , MODEL, Color, SAE  from  CRIS_INQUIRY  ")
    End If

    


    
    

        
        
       

    
    lvInquiry.Columns.DeleteAll
    i = 0
    
    If Not (temprs.EOF Or temprs.BOF) Then                    ''''COLUMN HEADERS
        For Each FLD In temprs.Fields
            lvInquiry.Columns.Add i, FLD.Name, 100, True
            lvInquiry.Columns(i).AutoSize = True
            lvInquiry.Columns(i).DrawFooterDivider = False
            i = i + 1
        Next
    End If
    If lvInquiry.Columns.Count > 0 Then
        lvInquiry.Columns(2).FooterText = "F3: Add Filter"
        lvInquiry.Columns(3).FooterText = "F8: Remove Filter"
    End If
    flex_FillReportView temprs, lvInquiry, False
End Sub

Private Sub cmdNewProspect_Click()
    frmCRIS_EntryProfilePersonal.Show
End Sub

Private Sub cmdOptionCal_Click()
    If cCalSales.ActiveView.GetSelectedEvents.Count = 0 Then: Exit Sub
    PopupMenu mnuContextCal

End Sub

Private Sub cmdP1_Click()
    Dim frm As New frmCRIS_EntryProspects
    frm.AddNewProspect
    frm.Show
End Sub

Private Sub cmdPivotVehicles_Click()
    Dim temprs                               As ADODB.Recordset
    Dim i                                    As Long
    Dim SQL                                  As String
    If cboMSAYear.Text = vbNullString Then
        MessagePop InfoVoid, "Select Year", "Select Year From The List", 1000, 1
        Exit Sub
    End If
    With lvInquiry
        .Columns.DeleteAll
        .Columns.Add 0, "Item", 50, True
        .Columns.Add 1, "SAE NAME", 350, True
        .Columns.Add 2, "JAN", 90, True
        .Columns.Add 3, "FEB", 90, True
        .Columns.Add 4, "MARCH", 90, True
        .Columns.Add 5, "APRIL", 90, True
        .Columns.Add 6, "MAY", 90, True
        .Columns.Add 7, "JUN", 90, True
        .Columns.Add 8, "JUL", 90, True
        .Columns.Add 9, "AUG", 90, True
        .Columns.Add 10, "SEP", 90, True
        .Columns.Add 11, "OCT", 90, True
        .Columns.Add 12, "NOV", 90, True
        .Columns.Add 13, "DEC", 90, True
        .Columns.Add 14, "SAE", 50, True
        .Columns(14).Visible = False
        .ShowGroupBox = False
    End With

    SQL = " SELECT (SElect Firstname + ' ' + LEFT(Middlename,1) + ' ' + Lastname  FROM HRMS_EMPINFO where ID = SAE) as SAE,  " _
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

    If optAdvSearch(2).Value = True Then
        SQL = Replace(SQL, "@APTYPE", 1)
        lblCap.caption = "Total Test Drive Made By SAE(s) By Months For The Year " & cboMSAYear.Text
    Else
        SQL = Replace(SQL, "@APTYPE", 2)
        lblCap.caption = "Total Vehicles Sales Appointments By SAE(s) By Months For The Year " & cboMSAYear.Text
    End If
    Set temprs = gconDMIS.Execute(SQL)



    flex_FillReportView temprs, lvInquiry, True
    Dim ilng(11)                             As Long

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
        ilng(11) = ilng(11) + CLng(lvInquiry.Rows(i).Record(13).Value)

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
    lvInquiry.Columns(13).FooterText = ilng(11)

    Erase ilng
End Sub


Private Sub cmdPrintLvInq_Click(Index As Integer)
    With lvInquiry
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End With
        lvInquiry.PrintOptions.BlackWhiteContrast = 0
        lvInquiry.PrintOptions.BlackWhitePrinting = True
        lvInquiry.PrintOptions.Header.Font.Size = "18"
        lvInquiry.PrintOptions.Header.Font.Underline = True
        lvInquiry.PrintOptions.Header.TextCenter = ReportTitle
            lvInquiry.PrintPreview True
    With lvInquiry
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
    End With
End Sub





Private Sub Command1_Click()
    Dim temprs                               As ADODB.Recordset
    Dim i                                    As Integer
    Dim SQL                                  As String
    Dim WhereIn                              As String


    SQL = "SELECT  + '(' + cast(Month(StartDate) as varchar)  + ')' + dbo.NameOfMonth(startDate) as MonthName ,  " _
        & " StartDate +  '-' + EndDate as Dates , " & _
          "StartTime + '-' + EndTime as [Time],  SAE, " & _
          "ProfileName  + '(' + ProspectType + ')' as [ClientName(Type)] , Model " & _
          "frOM CRIS_vW_Appointments Where AppointmentType=@APTYPE "




'SELECT     ProspectID, AcctName, LeadSource, VehicleCode, VehicleModel, Color,
'(SElect Firstname + ' ' + LEFT(Middlename,1) + ' ' + Lastname  FROM HRMS_EMPINFO where ID = SAE) as SAE
'From CRIS_Prospects



    For i = 1 To lvlist(0).ListItems.Count
        If lvlist(0).ListItems(i).Checked = True Then
            WhereIn = WhereIn & lvlist(0).ListItems(i).ListSubItems(1).Text & ", "
        End If
    Next
    If chkInqSAEIMonths(0).Value = 1 And chkInqSAEIYear(0).Value = 1 Then
        If Len(WhereIn) = 0 Then: MessagePop InfoVoid, "Invalid Selection ", "Select At Least A Month", 1000, 1: Exit Sub
        SQL = SQL & " and Month(StartDate) IN (" & Left(WhereIn, Len(WhereIn) - 2) & ") AND Year(StartDate)=" & cboInqSAEIYear(0).Text & " Order By 1,4"
    ElseIf chkInqSAEIMonths(0).Value = 1 And chkInqSAEIYear(0).Value = 0 Then
        If Len(WhereIn) = 0 Then: MessagePop InfoVoid, "Invalid Selection ", "Select At Least A Month", 1000, 1: Exit Sub
        SQL = SQL & " and  Month(StartDate) IN (" & Left(WhereIn, Len(WhereIn) - 2) & ") Order By 1,4"
    ElseIf chkInqSAEIMonths(0).Value = 0 And chkInqSAEIYear(0).Value = 1 Then
        SQL = SQL & "and  Year(StartDate)=" & cboInqSAEIYear(0).Text & " Order By 1,4"
    End If

    If optAdvSearch(1).Value = True Then
        SQL = Replace(SQL, "@APTYPE", 2)
    Else
        SQL = Replace(SQL, "@APTYPE", 1)
    End If



    Set temprs = gconDMIS.Execute(SQL)
    With lvInquiry
        .Columns.DeleteAll
        .Columns.Add 0, "Month", 0, True

        .Columns.Add 1, "Date", 150, True
        .Columns.Add 2, "Time", 120, True
        .Columns.Add 3, "SAE", 120, True
        .Columns.Add 4, "ClientName(Type)", 300, True
        .Columns.Add 5, "Model", 200, True
        .GroupsOrder.Add .Columns(0)
        .Columns(0).Visible = False
    End With
    For i = 0 To lvInquiry.Columns.Count - 1
        lvInquiry.Columns(i).DrawFooterDivider = False
    Next

    flex_FillReportView temprs, lvInquiry, False

    lvInquiry.Columns(1).FooterText = "F3: Add Filter"
    lvInquiry.Columns(2).FooterText = "F8: Remove Filter"

End Sub




'


'''''''''''''' ALL ABOUT PAGING



Friend Sub ConfigRecordSet(strQuery As String, GroupingHeader As Integer)
    Dim nCol                                 As ReportColumn
    Dim FLD                                  As Field
    Dim ijx                                  As Long

    Connect
    Set GridRs = New ADODB.Recordset
    With GridRs
        .CursorLocation = adUseClient
        .CacheSize = 5
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
        .Open strQuery, gconDMIS
        If GroupingHeader = 0 Then
            .PageSize = 50
        End If

    End With

    lvGrid.Columns.DeleteAll
    For Each FLD In GridRs.Fields
        Set nCol = lvGrid.Columns.Add(ijx, CStr(FLD.Name), 100, True)
        ijx = ijx + 1
    Next
    lvGrid.Columns(0).Visible = False

    If GridRs.RecordCount > 0 Then
        If GroupingHeader > 0 Then
            '   Paging rec_first
            Paging rec_all
            lvGrid.ShowGroupBox = True
            lvGrid.GroupsOrder.Add lvGrid.Columns(GroupingHeader)
            lvGrid.Columns(GroupingHeader).Visible = False
        Else
            lvGrid.ShowGroupBox = False
            Paging rec_first
            'Paging rec_all
        End If

    Else
        lvGrid.ShowGroupBox = False
        lvGrid.Records.DeleteAll

    End If
    lvGrid.Populate
End Sub



Private Sub Form_Load()
    InitVars
    wGraphs.Navigate ("about:<body scroll='no' ></body>")
    optAdvSearch_Click (0)
    cmdInqSAEISales_Click
End Sub

Private Sub ImgSearchProspect_Click()
    Dim temprs                               As ADODB.Recordset
    Dim SQL                                  As String
    SQL = "SELECT ProfileID,convert(varchar , InquiryDate ,101) , AcctName, (Select masterdata from CRIS_vw_MASTER_pULLDOWN  WHERE  DataID=LeadSource), VehicleModel,COlor ,SAE FROM  CRIS_Profile "

    If cboFilter.Text = "Phone" Then
        SQL = SQL & " WHERE BusinessPhone in (@RR) or  OtherPhone in (@RR)  or CellPhone in (@RR)  or Fax in (@RR) "
        SQL = Replace(SQL, "@RR", "'" & txtsearch.Text & "'")
    Else
        SQL = SQL & " WHERE Replace(" & cboFilter.Text & ",',','') Like '%" & Replace(txtsearch.Text, "'", "") & "%' "
    End If
    
    If chkProspectInq(0).Value = 1 Then: SQL = SQL & " AND VehicleModel='" & cboFilterInq(0).Text & "'"
    If chkProspectInq(1).Value = 1 Then: SQL = SQL & " AND color='" & cboFilterInq(1).Text & "'"
    If chkProspectInq(2).Value = 1 Then: SQL = SQL & " AND SAE='" & cboFilterInq(2).Text & "'"
    If chkProspectInq(3).Value = 1 Then: SQL = SQL & " AND LeadSource=" & cboFilterInq(3).ItemData(cboFilterInq(3).ListIndex)


    Set temprs = gconDMIS.Execute(SQL & "order by 3")

    flex_FillReportView temprs, lvGrid, False


    lvGrid.Populate

    If lvGrid.Records.Count = 0 Then
        lblProspectHitList.ForeColor = vbRed
    Else
        lblProspectHitList.ForeColor = &H853036
    End If




    lblProspectHitList.caption = lvGrid.Records.Count & " Records Found"
    frmCRIS_PaneContacts.picCustList.Visible = IIf(lvGrid.Records.Count = 0, False, True)

End Sub

Sub InitVars()


    With lvInquiry                                            '''''''''''UI
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        '.PaintManager.ShadeSortColor = RGB(209, 209, 209)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
    End With








    Call ComboList(cboInqSAEISales(0), gconDMIS.Execute("Select DISTINCT 1 , SAE from CRIS_INQUIRY "))
    Call ComboList(cboInqSAEISales(1), gconDMIS.Execute("SELECT  DISTINCT  1, Color  FROM CRIS_INQUIRY Order by 2 "))
    Call ComboList(cboInqSAEISales(2), gconDMIS.Execute("SELECT  DISTINCT  1, MODEL FROM CRIS_INQUIRY Order by 2"))
    Call ComboList(cboInqSAEISales(3), gconDMIS.Execute("SELECT  DISTINCT 1, LeadSource  FROM CRIS_INQUIRY Order by 2 "))

    
    
    Dim i                                    As Integer
    Dim j                                    As Integer
'    For i = 1 To 12
'            lvINQList.ListItems.Add , , MonthName(i)
'            lvINQList.ListItems(i).ListSubItems.Add , , i
'
'            lvINQList1.ListItems.Add , , MonthName(i)
'            lvINQList1.ListItems(i).ListSubItems.Add , , i
'
'    Next
'            lvINQList.ColumnHeaders(1).Width = lvINQList.Width * 0.85
'            lvINQList1.ColumnHeaders(1).Width = lvINQList1.Width * 0.85
'
    For i = 0 To 5
    cboMSAYear.AddItem Year(Now) - i
'        cboInqSAEISales(4).AddItem Year(Now) - i
'        'cboInqSAEISales(0).AddItem Year(Now) - i
    Next
    cboMSAYear.ListIndex = 0
'        cboInqSAEISales(4).ListIndex = 0
''        cboINQYear1.ListIndex = 0
End Sub




Private Sub lblImgSearch_Click()
ImgSearchProspect_Click
End Sub

Private Sub lvGrid_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record Is Nothing Then Exit Sub
    If Row.Record(1).Value = "PP" Or Row.Record(1).Value = "PC" Then
        Metrics.BackColor = RGB(234, 250, 218)

    Else
    End If
        
End Sub

Private Sub lvGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lvGrid.Records.Count > 0 Then
        Call frmCRIS_Filter.ConfigGrid(lvGrid, 3)
        frmCRIS_Filter.Show vbModeless
    ElseIf KeyCode = vbKeyF8 And lvGrid.Records.Count > 0 Then
        lvGrid.FilterText = vbNullString
        lvGrid.Populate
        lvGrid.Columns(4).FooterText = vbNullString
    ElseIf KeyCode = vbKeyUp And lvGrid.SelectedRows(0).Index = 0 Then
        txtsearch.SetFocus
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    lvGrid_RowRClick Row, Item
End Sub

Private Sub lvGrid_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mnu1(6).Enabled = IIf(IsNull(lvGrid.SelectedRows.Row(0).Record(5).Value) = True, False, True)
    mnu1(7).Enabled = Not (mnu1(6).Enabled)
    mnu1(8).Enabled = Not (mnu1(7).Enabled)
    PopupMenu mnuContextGrid
End Sub

Private Sub lvGrid_SelectionChanged()
    If lvGrid.Records(0) Is Nothing Then: Exit Sub
    Dim Qstr  As String
        Dim temprs                           As ADODB.Recordset
'
Dim SQL As String
SQL = "SELECT  Salutations, FirstName, LastName, MI, DateofBirth, Sex, SpouseName, Anniversary, JobTitle, CompanyName, IndustryType, isCompany, ContactType, CustomerClassification, Comp_Address, Res_Address, Ship_Address, Billing_Address, HomePhone, BusinessPhone, OtherPhone, Cellphone, Fax, Email, AssistantName, AsstPhone," _
& " (select MasterData from CRIS_VW_MASTER_PULLDOWN WHERE DATAID= CustomerClassification) as classification, " _
& " convert(varchar, InquiryDate ,101) InquiryDate,  convert(varchar, LogQuote ,101) LogQuote , convert(varchar, LogEmail ,101) LogEmail, convert(varchar, LogAppointment ,101) LogAppointment, " _
& " convert(varchar, LogTestDrive ,101) LogTestDrive,  convert(varchar, LogCall ,101) LogCall ,  convert(varchar, LogJournal ,101) LogJournal , convert(varchar, LogLetter ,101)  LogLetter  FROM CRIS_Profile WHERE PROFILEID=" & lvGrid.SelectedRows(0).Record(0).Value

            Set temprs = gconDMIS.Execute(SQL)
            If Not (temprs.EOF Or temprs.BOF) Then
                With frmCRIS_PaneContacts
                    .lblProspectProfile = Null2String(temprs("LastName"))
                    .lblProspectProfile = .lblProspectProfile & DateDiff("D", temprs("Dateofbirth"), Now)
                    .lblProspectProfile = .lblProspectProfile & Null2String(temprs("SpouseName"))
                    .lblProspectClassification = Null2String(temprs("classification"))
                    If IsNull(temprs("LogQuote")) = False Then: Qstr = Chr(149) & "Quotation Sent (" & temprs("LogQuote") & ")" & vbCrLf
                    If IsNull(temprs("LogEmail")) = False Then: Qstr = Qstr & Chr(149) & "Email Sent (" & temprs("LogEmail") & ")" & vbCrLf
                    If IsNull(temprs("LogAppointment")) = False Then: Qstr = Qstr & Chr(149) & "Appointment Made (" & temprs("LogAppointment") & ")" & vbCrLf
                    If IsNull(temprs("LogTestDrive")) = False Then: Qstr = Qstr & Chr(149) & "Test Drive Scheduled (" & temprs("LogTestDrive") & ")" & vbCrLf
                    If IsNull(temprs("LogCall")) = False Then: Qstr = Qstr & Chr(149) & "Calls Made (" & temprs("LogCall") & ")" & vbCrLf
                    If IsNull(temprs("LogJournal")) = False Then: Qstr = Qstr & Chr(149) & "Journals Added (" & temprs("LogJournal") & ")" & vbCrLf
                    If IsNull(temprs("LogLetter")) = False Then: Qstr = Qstr & Chr(149) & "Letter Sent (" & temprs("LogLetter") & ")" & vbCrLf
                    
                            .lblProspectStatus = Qstr
                End With
'
            End If
'        Else
'            Set temprs = gconDMIS.Execute("Select AcctName, " _
'                                       & " (Select MasterData from  CRIS_vw_Master_PullDown where MasterType='Lead Source' and DataID=LeadSource) as LeadSource,  " _
'                                       & " DateofBirth,Sex, SpouseName,CompanyName,  " _
'                                       & " (Select MasterData from  CRIS_vw_Master_PullDown where MasterType='Customer Classification' and DataID=LeadSource) as CustomerClassification,  " _
'                                       & " PrimaryEmail,  PrimaryContact From CRIS_PROFILE  " _
'                                       & " Where ProfileId = " & lvGrid.SelectedRows(0).Record(0).Value)
'
'            If Not (temprs.EOF Or temprs.BOF) Then
'                With frmCRIS_PaneContacts
'                    .lblCustomerName = Null2String(temprs("AcctName"))
'                    .lblLeadSource = Null2String(temprs("LeadSource"))
'                    If IsNull(temprs("DateofBirth")) = False Then
'                        .lblAgeSex.caption = DateDiff("yyyy", temprs("DateofBirth"), Now) & "/" & Null2String(temprs("sex"))
'                    Else
'                        .lblAgeSex = Null2String(temprs("DateOfbirth")) & "/" & Null2String(temprs("sex"))
'                    End If
'                    .lblSpouse = Null2String(temprs("SpouseName"))
'                    .lblCompany = Null2String(temprs("CompanyName"))
'                    .lblClassification = Null2String(temprs("CustomerClassification"))
'                    .lblEmail = Null2String(temprs("PrimaryEmail"))
'                    .lblPhone = Null2String(temprs("PrimaryContact"))
'                End With
'
'            End If
'        End If
'
'    End If
End Sub

'Private Sub lvInquiry_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
'    If Row.Record Is Nothing Then Exit Sub
'    'If optAdvSearch(3).Value = True Or optAdvSearch(4).Value = True Then
'
'     '   If Item.Index > 1 Then
'     '       If Item.Value > 0 Then
'     '           Metrics.ForeColor = vbBlue
'     '           Metrics.Font.Bold = True
'     '           Metrics.BackColor = RGB(189, 202, 236)
'     '       Else
'     '           Metrics.ForeColor = vbRed
'     '           Metrics.Font.Strikethrough = True'
'
'            'End If
'        'End If
'
'    'End If
'End Sub

Private Sub lvInquiry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lvInquiry.Records.Count > 0 Then
        
        Call frmCRIS_Filter.ConfigGrid(lvInquiry, 3)
        frmCRIS_Filter.Show vbModeless
    ElseIf KeyCode = vbKeyF8 And lvInquiry.Records.Count > 0 Then
        lvInquiry.FilterText = vbNullString
        lvInquiry.Populate
        lvInquiry.Columns(4).FooterText = vbNullString
    End If
End Sub

Private Sub lvInquiry_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
        If Row.Record Is Nothing Then: Exit Sub
        If optAdvSearch(0).Value = True Then
        Call frmCRIS_ViewLog.ShowReport(Row.Record(0).Value, Row.Record(1).Value)
            frmCRIS_ViewLog.Show
        ElseIf optAdvSearch(0).Value = True Then
        
        ElseIf optAdvSearch(2).Value = True Then
         Call frmCRIS_ViewLog.ShowSAEAppointmentDetail(Row.Record(14).Value, Row.Record(1).Value)
            frmCRIS_ViewLog.Show
        End If
        


        
End Sub

Private Sub lvInquiry_SelectionChanged()
'    On Error GoTo adder:
'    If lvInquiry.SelectedRows.Count <= 0 Then: Exit Sub
'    If Not (optAdvSearch(3).Value = True Or optAdvSearch(4).Value = True) Then: Exit Sub
'    Screen.MousePointer = 11
'    Dim strQ(), i, WordJavascript, j, MyVal
'    If lvInquiry.Records.Count <= 0 Then: Exit Sub
'
'    ReDim strQ(lvInquiry.SelectedRows.Count - 1)
'
'    For j = 0 To lvInquiry.SelectedRows.Count - 1
'        For i = 2 To 13
'            strQ(j) = strQ(j) & lvInquiry.SelectedRows.Row(j).Record(i).Value & ", "
'        Next
'    Next
'
'
'    WordJavascript = "<html><head>" & vbCrLf & _
'                     "<SCRIPT LANGUAGE='JavaScript1.2' SRC='graph.js'></SCRIPT></head> " & vbCrLf & _
'                     "<body oncontextmenu='return false;' style='margin:5 0 0 0; overflow:hidden;border:0px solid white;'>" & vbCrLf & _
'                     "<SCRIPT LANGUAGE='JavaScript'> " & vbCrLf & _
'                     "var q = new Graph(260,105); " & vbCrLf & _
'                     "q.scale =10; " & vbCrLf & _
'                     "q.setXScale('',1); " & vbCrLf & _
'                     "@QSTRING" & vbCrLf & _
'                     "q.build(); " & vbCrLf & _
'                     "</SCRIPT> " & vbCrLf & _
'                     "</body>"
'    ' "q.title = 'Appointments'; " & vbCrLf & _
'      "q.yLabel = '#App'; " & vbCrLf & _
'      "q.xLabel = 'Date'; " & vbCrLf & _
'
'      For i = 0 To UBound(strQ)
'    strQ(i) = "q.addRow(" & MID(strQ(i), 1, Len(strQ(i)) - 2) & ");"
'Next
'
'For i = 0 To UBound(strQ)
'    MyVal = MyVal & strQ(i) & vbCrLf
'Next
'
'
'
'WordJavascript = Replace(WordJavascript, "@QSTRING", MyVal)
'
'
'Dim fso                                      As New FileSystemObject
'
'Dim stream                                   As TextStream
'Set stream = fso.OpenTextFile(App.Path & "/graphs/graph.html", ForWriting, True)
'stream.WriteLine (WordJavascript)
'stream.Close
'Set stream = Nothing
'Set fso = Nothing
'
'wGraphs.Navigate (App.Path & "/graphs/graph.html")
'Screen.MousePointer = 0
'Exit Sub
'adder:
'Err.Clear
'MessagePop InfoWarning, "Path Acess errror", "Check You Path Setting "
End Sub

Private Sub mnu1_Click(Index As Integer)
    Dim CustomerType                         As String
    Dim CustomerID                           As Long
    On Error GoTo adder:
    CustomerType = lvGrid.SelectedRows.Row(0).Record(1).Value
    CustomerID = lvGrid.SelectedRows.Row(0).Record(0).Value

    Select Case Index
        Case 0                                                '0   Add sales
            Call frmCRIS_EntrySalesAppointment.NewEvent(CustomerID, CustomerType)
            frmCRIS_EntrySalesAppointment.Show
        Case 1                                                '1   testdrive

            Call frmCRIS_EntryTestDriveAppointment.NewEvent(CustomerID, CustomerType)
            frmCRIS_EntryTestDriveAppointment.Show
        Case 3                                                '3   quotation
            Call frmCRIS_EntryQuotation.NewQuotation(CustomerID, CustomerType)
            frmCRIS_EntryQuotation.Show vbModal
        Case 4                                                '4   invitation


        Case 6 To 7
            '6 change, 7 add group
            Set FormGroup = New frmCRIS_Group
            With lvGrid.SelectedRows.Row(0)
                FormGroup.CustomerID = .Record(0).Value
                FormGroup.CustomerType = .Record(1).Value
                FormGroup.DataID = IIf(IsNull(.Record(5).Value) = True, 0, .Record(5).Value)
                FormGroup.Show vbModal
            End With
        Case 8                                                '8   Remove Group

            If MsgBox("Confirm Your Action", vbOKCancel Or vbQuestion Or vbDefaultButton1, App.Title) = vbCancel Then: Exit Sub
            If CustomerType = "CC" Or CustomerType = "CP" Then
                gconDMIS.Execute ("Update CRIS_PROFILE Set ContactType=NULL Where ProfileID= " & CustomerID)
            Else
                gconDMIS.Execute ("Update ALL_CUSTOMER Set ContactType=NULL Where ID= " & CustomerID)
            End If
            MessagePop DELETE, " Removed", "Profile Group Removed"
            lvGrid.SelectedRows.Row(0).Record(5).Value = Null
        Case 10                                               ' Log A Call
            With lvGrid.SelectedRows.Row(0)
                Call frmCRIS_LogCall.AddCall(CustomerID, CustomerType, .Record(3).Value & vbCrLf & .Record(4).Value)
            End With
            frmCRIS_LogCall.Show
        Case 11                                               ' Log A Journal

            With lvGrid.SelectedRows.Row(0)
                Call frmCRIS_LogJournal.AddJournal(CustomerID, CustomerType, .Record(3).Value & vbCrLf & .Record(4).Value)
            End With
            frmCRIS_LogJournal.Show
        Case 13

            If CustomerType = "CC" Or CustomerType = "CP" Then
                MsgBox "Customer Information Editing Not Yet Implemented"
            Else
                frmCRIS_EntryProfilePersonal.ProfileID = CustomerID
                frmCRIS_EntryProfilePersonal.Show
            End If

        Case 15                                               ' Print
            lvGrid.PrintOptions.BlackWhitePrinting = True
            lvGrid.PrintOptions.Header.TextCenter = "Customer Listing"
            lvGrid.PrintPreview True
    End Select
    Exit Sub
adder:
    Err.Clear

End Sub

Private Sub mnuAppointments_Click(Index As Integer)
    If Index = 0 Then
        frmCRIS_EntrySalesAppointment.NewEvent 0, 0
        frmCRIS_EntrySalesAppointment.Show
        Set frmCRIS_EntrySalesAppointment = Nothing
    Else
        Call frmCRIS_EntryTestDriveAppointment.NewEvent(0, 0)
        frmCRIS_EntryTestDriveAppointment.Show
        Set frmCRIS_EntryTestDriveAppointment = Nothing
    End If
End Sub
'''''''''END MENU LINES

Private Sub mnuCalOptionAllAppointments_Click()
    If cCalSales.ActiveView.GetSelectedEvents.Count <= 0 Then: Exit Sub
    Dim viewX                                As CalendarEvent
    Dim i                                    As Integer
    Dim ar                                   As String
    Dim f                                    As frmCRIS_ViewAllApointments
    For i = 0 To cCalSales.ActiveView.GetSelectedEvents.Count - 1
        Set viewX = cCalSales.ActiveView.GetSelectedEvents(i).Event
        ar = ar & viewX.ID & ","
    Next
    ar = Mid(ar, 1, Len(ar) - 1)
    Set f = New frmCRIS_ViewAllApointments

    If ShortcutBar.Selected.ID = 0 Then
        Call f.SQLString(ar, 2)
    Else
        Call f.SQLString(ar, 1)
    End If
    f.Show
    Set viewX = Nothing
End Sub

Private Sub mnuDeleteEvent_Click()
    If (MsgBox("Are You Sure ", vbOKCancel)) = vbOK Then
        cCalSales.DataProvider.DeleteEvent ContextEvent
        cCalSales.Populate
    End If
End Sub

Private Sub mnuOpenEvent_Click()
    If ShortcutBar.Selected.ID = 2 Then
        ModifyEvent ContextEvent, 1
    Else

        ModifyEvent ContextEvent, 2
    End If
End Sub

'''''''''MENU LINES
Private Sub mnuTimeScale_Click(Index As Integer)
    cCalSales.DayView.TimeScale = mnuTimeScale(Index).HelpContextID
    cCalSales.DayView.ScrollToWorkDayBegin
    cCalSales.RedrawControl

End Sub

Private Sub ModifyEvent(ModEvent As CalendarEvent, nit As Integer)
    If nit = 1 Then
        frmCRIS_EntrySalesAppointment.ModifyEvent ModEvent
        frmCRIS_EntrySalesAppointment.Show vbModal
    Else
        frmCRIS_EntryTestDriveAppointment.ModifyEvent ModEvent
        frmCRIS_EntryTestDriveAppointment.Show vbModal
    End If
End Sub

Public Sub OpenProvider(ByVal strConnectionString As String, nID As Integer)
    Set m_pCustomDataHandler = Nothing
    Set m_pCustomDataHandler = New providerSQLServer
    m_pCustomDataHandler.OpenDB strConnectionString
    m_pCustomDataHandler.SetCalendar cCalSales, nID
    cCalSales.SetDataProvider strConnectionString
    cCalSales.DataProvider.CacheMode = xtpCalendarDPCacheModeOnRepeat
    If Not cCalSales.DataProvider.Open Then
        cCalSales.DataProvider.Create
    End If
    cCalSales.Populate
    dtCal.RedrawControl
End Sub

Private Sub optAdvSearch_Click(Index As Integer)
    Dim i                                    As Integer
    
    For i = 0 To picInq.Count - 1
        picInq(i).Visible = False
    Next
    
    lvInquiry.Columns.DeleteAll
    lvInquiry.Records.DeleteAll
    lvInquiry.Populate
    
    Select Case Index
        Case 0
            If optAdvSearch(Index).Value = True Then: picInq(Index).Visible = True
            ReportTitle = "Prospect Inquiry"
            cmdInqSAEISales_Click
        Case 1
            If optAdvSearch(Index).Value = True Then: picInq(Index).Visible = True
            ReportTitle = "Monthly Sales Appointment"
        Case 2
            If optAdvSearch(Index).Value = True Then: picInq(Index).Visible = True
           
             ReportTitle = "Monthly Sale Appointments By SAE"
            cmdPivotVehicles_Click
        Case 3

            


        Case 4
            

        Case 5

    End Select
        Me.caption = "Inquiry :" & ReportTitle

End Sub

Friend Sub Paging(PageTo As MoveWhere)
    Dim FLD                                  As Field
    Dim PageCount                            As Integer
    Dim AbsPage                              As Integer
    Dim REC                                  As ReportRecord
    Dim j                                    As Long
    PageCount = GridRs.PageCount
    lvGrid.Records.DeleteAll
    lvGrid.FilterText = vbNullString
    If PageCount = 0 Then: lvGrid.Populate: Exit Sub
    AbsPage = GridRs.AbsolutePage

    Select Case PageTo
        Case rec_next

        Case rec_prev
            If AbsPage = 2 Then
                GridRs.AbsolutePage = 1
            Else
                GridRs.AbsolutePage = AbsPage - 2
            End If
        Case rec_first
            GridRs.AbsolutePage = 1
        Case rec_last
            GridRs.AbsolutePage = GridRs.PageCount
        Case rec_all

            While Not GridRs.EOF
                Set REC = lvGrid.Records.Add
                For Each FLD In GridRs.Fields
                    REC.AddItem (Trim(FLD.Value))
                Next
                GridRs.MoveNext
            Wend

            lvGrid.Populate
            Exit Sub
    End Select
    ''''''''Iterate
    For j = 0 To GridRs.PageSize
        If GridRs.EOF Then: Exit For
        Set REC = lvGrid.Records.Add
        For Each FLD In GridRs.Fields
            REC.AddItem (Trim(FLD.Value))
        Next
        GridRs.MoveNext
    Next

    ''''''''get it
    AbsPage = GridRs.AbsolutePage
    ''''''''pageit
    Select Case AbsPage
        Case 2:

        Case -2:

        Case -3:
            If PageCount = 1 Then

            Else

            End If
        Case PageCount

        Case Else

    End Select
    ''''''''handle position
    If AbsPage = adPosEOF Then
        GridRs.AbsolutePage = GridRs.PageCount
    End If
    ''''''''throw

    lvGrid.Populate

End Sub

Private Sub picInq_Resize(Index As Integer)
    If Index = 2 Then
        wGraphs.Left = picInq(Index).ScaleWidth - wGraphs.Width
    End If
End Sub

Private Sub picInquiry_Resize()
    Dim i                                    As Integer
    For i = 0 To picInq.Count - 1
        picInq(i).Move 0, 0, picInquiry.ScaleWidth
    Next
    lvInquiry.Left = 0
    lvInquiry.Top = picInq(0).Height + 10
    lvInquiry.Height = picInquiry.ScaleHeight - (picInq(0).Top + picInq(0).ScaleHeight)
    lvInquiry.Width = picInquiry.ScaleWidth

End Sub



Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If
End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .LoadDesignerBars
        .LoadCommandBars "CRIS", App.Title, "Layout"
        '  .PaintManager.ClearTypeTextQuality = True

    End With
    '    With SkinFramework1
    '        .LoadSkin App.Path & "/r.html", ""
    '        .ApplyWindow Me.hwnd
    '        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    '    End With

    '    Dim ToolTipContext                       As ToolTipContext
    '    Set ToolTipContext = CommandBars1.ToolTipContext
    '    With ToolTipContext
    '        .ShowTitleAndDescription True, xtpToolTipIconInfo
    '        .SetMargin 2, 2, 2, 2
    '        .MaxTipWidth = 180
    '        If .IsBalloonStyleSupported Then
    '            .Style = xtpToolTipBalloon
    '        Else
    '            .Style = xtpToolTipOffice2007
    '        End If
    '        .ShowShadow = True
    '    End With
End Sub


Private Sub PopCntrl_StateChanged()
    '    If PopCntrl.State = xtpPopupStateClosed Then
    '        Dim frm                              As Form
    '        For Each frm In Forms
    '            frm.Enabled = True
    '        Next
    '    End If
End Sub

Sub SetCaptionsVisualTheme()
    On Error Resume Next
    Dim CtrlCaption                          As ShortcutCaption
    Dim Form As Form, Ctrl                   As Object

    For Each Form In Forms
        For Each Ctrl In Form.ControlS
            Set CtrlCaption = Ctrl
            If Not CtrlCaption Is Nothing Then
                CtrlCaption.VisualTheme = ShortcutBar.VisualTheme
            End If
        Next
    Next
End Sub

Private Sub shortcutbar_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
    captionTitle.caption = Item.caption
    Dim i                                    As Integer
    Dim temprs                               As ADODB.Recordset
    Dim SQL                                  As String
    Select Case Item.ID
        Case 0
            PictureCustomers.Visible = False
            picInquiry.Visible = False
            cCalSales.Visible = True
            picMainDash.Visible = False
            OpenProvider "Provider=Custom;" & SQLConnectionString, 2
        Case 1
            PictureCustomers.Visible = False
            picInquiry.Visible = False
            cCalSales.Visible = True
            picMainDash.Visible = False
            OpenProvider "Provider=Custom;" & SQLConnectionString, 1
        Case 2
            cCalSales.Visible = False
            PictureCustomers.Visible = True
            picInquiry.Visible = False
picMainDash.Visible = False

'ProfileID, ProfileType, AcctName, ProfileName, Address ,ContactType

            'SQL = "SELECT pROSPEF,convert(varchar , InquiryDate ,101) , AcctName, (Select masterdata from CRIS_vw_MASTER_pULLDOWN  WHERE  DataID=LeadSource), VehicleModel,COlor ,SAE FROM  CRIS_Profile "
        
            'Set temprs = gconDMIS.Execute(SQL)
            'flex_FillReportView temprs, lvGrid, False
            'lvGrid.Populate

        Case 3
            cCalSales.Visible = False
            PictureCustomers.Visible = False
            picInquiry.Visible = True
            picMainDash.Visible = False
        Case 4
            cCalSales.Visible = False
            PictureCustomers.Visible = False
            picInquiry.Visible = False '
            picMainDash.Visible = True
    End Select






End Sub

Private Sub ThemeCalendar(cCal As CalendarControl, picKer As DatePicker)
    Dim oThemeDT                             As DatePickerThemeOffice2007

    With cCal

        .EnableReminders False
        .Options.EnableInPlaceCreateEvent = False
        .Options.EnableInPlaceEditEventSubject_AfterEventResize = False
        .Options.EnableInPlaceEditEventSubject_ByF2 = False
        .Options.EnableInPlaceEditEventSubject_ByMouseClick = False
        .Options.EnableInPlaceEditEventSubject_ByTab = False
        .Options.DayViewTimeScaleShowMinutes = True
        .Options.DayViewCurrentTimeMarkVisible = 1
        .Options.MonthViewCompressWeekendDays = True
        .Options.MonthViewShowEndDate = True
        .Options.WeekViewShowEndDate = True
        .Options.WorkWeekMask = xtpCalendarDayMo_Fr
        .ViewType = xtpCalendarDayView
        .EnableToolTips (False)
        
    End With
    picKer.AttachToCalendar cCal
    picKer.SetTheme oThemeDT
    picKer.RedrawControl
    
    cCal.DayView.ScrollToWorkDayBegin
    cCal.Populate
    cCal.RedrawControl
End Sub

Private Sub txtsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lvGrid.Records.Count > 0 Then
            lvGrid.Rows(0).Selected = True
            lvGrid.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ImgSearchProspect_Click
    End If
End Sub

