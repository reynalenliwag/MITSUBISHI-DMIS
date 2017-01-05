VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_Visits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Visits"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogVisits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   7755
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5295
      ScaleWidth      =   2805
      TabIndex        =   0
      Top             =   1785
      Width           =   2805
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   465
         Width           =   2760
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4395
         Left            =   45
         TabIndex        =   2
         Top             =   885
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7752
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Search By Date"
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
         Height          =   225
         Left            =   120
         TabIndex        =   36
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7755
      TabIndex        =   23
      Top             =   0
      Width           =   7755
      Begin VB.TextBox txtEntityName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   210
         Width           =   4935
      End
      Begin VB.TextBox txtEntityContactperson 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtEntityAddress 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "LogVisits.frx":08CA
         Top             =   1230
         Width           =   4935
      End
      Begin VB.TextBox txtEntityPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox txtEntityMobile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   25
         Text            =   "09175041620"
         Top             =   720
         Width           =   2370
      End
      Begin VB.TextBox txtEntityEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   5070
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1260
         Width           =   2370
      End
      Begin VB.Label labEntityName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label labEntityAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   34
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CONTACT PERSON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   33
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   32
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   31
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   30
         Top             =   510
         Width           =   1230
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   6765
         Y1              =   1710
         Y2              =   1710
      End
   End
   Begin VB.PictureBox picDataEntry 
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   2835
      ScaleHeight     =   4395
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   1785
      Width           =   4695
      Begin VB.TextBox txtVisitComments 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   150
         TabIndex        =   18
         Top             =   1035
         Width           =   4395
      End
      Begin VB.TextBox txtVisitResults 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   150
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2580
         Width           =   4425
      End
      Begin MSComCtl2.DTPicker txtVisitDate 
         Height          =   345
         Left            =   150
         TabIndex        =   19
         Top             =   465
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39139
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
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
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   795
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date Visited"
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
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Results"
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
         Height          =   225
         Left            =   150
         TabIndex        =   20
         Top             =   2280
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7770
      TabIndex        =   4
      Top             =   6180
      Width           =   7770
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   6240
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   6
         Top             =   15
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "LogVisits.frx":0948
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":0A9A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "LogVisits.frx":0DD8
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":0F2A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Save this Record"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   2490
         ScaleHeight     =   915
         ScaleWidth      =   5655
         TabIndex        =   9
         Top             =   30
         Width           =   5655
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "LogVisits.frx":127A
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":13CC
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Exit Window"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "LogVisits.frx":1732
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":1884
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Delete Selected Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "LogVisits.frx":1BAF
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":1D01
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "LogVisits.frx":205D
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":21AF
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Add Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "LogVisits.frx":24C2
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2614
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Find a Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1080
            MouseIcon       =   "LogVisits.frx":290E
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2A60
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   390
            MouseIcon       =   "LogVisits.frx":2DB8
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2F0A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   5
         Top             =   75
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Visits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim ENTRY_LOGID                                                       As Long
Dim RS                                                                As ADODB.Recordset
Dim CustomerCode                                                      As String
Dim LOCALMODULENAME                                                   As String

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset

    ListView1.Enabled = False

    If CustomerCode <> vbNullString Then
        Set TEMPRS = gconDMIS.Execute("SELECT DATETIMEVISIT,COMMENTS , LOGID  FROM CRIS_PROSPECT_VISITS " & _
                                    " WHERE  CSCDE=" & N2Str2Null(CustomerCode) & " AND  CONVERT(VARCHAR, DATETIMEVISIT , 101)  LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY DATETIMEVISIT  ASC")
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT DATETIMEVISIT,COMMENTS , LOGID  FROM CRIS_PROSPECT_VISITS " & _
                                    " WHERE  PROSPECTID=" & PROSPECTID & " AND  CONVERT(VARCHAR, DATETIMEVISIT , 101)  LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY DATETIMEVISIT  ASC")
    End If

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        ListView1.Enabled = True
    End If

    flex_FillListView TEMPRS, ListView1



End Sub

Sub InitData()

    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False

    AddColumnHeader "Date , EmailAddress", ListView1
    ResizeColumnHeader ListView1, "40,55"
    FillSearchGrid ""


End Sub

Sub InitMemVars()
    txtVisitDate = DateValue(LOGDATE)
    txtVisitComments = ""
    txtVisitResults = ""


End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    If CustomerCode <> vbNullString Then
        RS.Open "SELECT * From CRIS_Prospect_Visits Where CSCDE=" & N2Str2Null(CustomerCode) & "Order BY DateTimeVisit desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        RS.Open "SELECT * From CRIS_Prospect_Visits Where ProspectID=" & PROSPECTID & "Order BY DateTimeVisit desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    txtEntityAddress = ""
    txtEntityContactperson = ""
    txtEntityEmail = ""
    txtEntityMobile = ""
    txtEntityName = ""
    txtEntityPhone = ""

    If xProspectID = 0 Then
        labEntityName = "CUSTOMER NAME"
        Set TEMPRS = gconDMIS.Execute("Select CUSTOMERNAME as [Name], CONTACTPERSON, PHONE, MOBILE, ADDRESS, EMAIL from CRIS_VW_ALLPROFILE WHERE CUSCDE=" & N2Str2Null(xCUSCODE))

    Else
        labEntityName = "PROSPECT NAME"
        Set TEMPRS = gconDMIS.Execute("Select ACCTNAME As [NAME], CONTACTPERSON, TELEPHONE as PHONE , MOBILE, ADDRESS , EMAIL  from CRIS_PROSPECTS WHERE PROSPECTID=" & N2Str2Null(xProspectID))

    End If

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtEntityAddress = Null2String(TEMPRS!Address)
        txtEntityContactperson = Null2String(TEMPRS!ContactPerson)
        txtEntityEmail = Null2String(TEMPRS!EMAIL)
        txtEntityMobile = Null2String(TEMPRS!Mobile)
        txtEntityName = Null2String(TEMPRS!Name)
        txtEntityPhone = Null2String(TEMPRS!Phone)
    End If
    Set TEMPRS = Nothing
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        'SELECT ENTRY_LOGID, ProspectID, DateEmail, EmailFrom, EmailTO, Subject, Body, Bound FROM DMIS.dbo.CRIS_Prospect_Email
        ENTRY_LOGID = RS!LOGID
        PROSPECTID = RS!PROSPECTID
        txtVisitComments = Null2String(RS!Comments)
        If IsNull(RS!DateTimeVisit) = False Then
            txtVisitDate = RS!DateTimeVisit
        End If
        txtVisitResults = Null2String(RS!Results)

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DateTimeVisit) FROM CRIS_Prospect_VISITS  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGVISIT=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGVISIT=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Friend Sub AddVisit(xProsID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    PROSPECTID = xProsID
    CustomerCode = xCustCode
    If xProsID = 0 Then
        LOCALMODULENAME = "CUSTOMER LOG"
    Else
        LOCALMODULENAME = "PROSPECT LOG"
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_add", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:

    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    txtVisitDate.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    ENTRY_LOGID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        'gconDMIS.Execute "delete from CRIS_Prospect_Visits where Logid=" & ENTRY_LOGID
        SQL_STATEMENT = "delete from CRIS_Prospect_Visits where Logid=" & ENTRY_LOGID
        gconDMIS.Execute (SQL_STATEMENT)

        NEW_LogAudit "X", "LOG PROSPECT VISIT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        UpdateLog
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If

    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    txtVisitDate.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    On Error GoTo ErrorCode:

    Unload Me





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdSave_Click()
    Dim SQL                                                           As String
    On Error GoTo ErrorCode:

    If ENTRY_LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospect_Visits (ProspectID, CSCDE, DateTimeVisit, Comments,Results) VALUES ("
        SQL = SQL & PROSPECTID & ", "
        SQL = SQL & N2Str2Null(CustomerCode) & ", "
        SQL = SQL & N2Str2Null(txtVisitDate.Value) & ", "
        SQL = SQL & N2Str2Null(txtVisitComments) & ", "
        SQL = SQL & N2Str2Null(txtVisitResults) & ")"
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "LOG PROSPECT VISIT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

    Else
        SQL = "Update CRIS_Prospect_Visits SET  "
        SQL = SQL & " DateTimeVisit=" & N2Str2Null(txtVisitDate.Value) & " , "
        SQL = SQL & " Comments=" & N2Str2Null(txtVisitComments) & " , "
        SQL = SQL & " Results=" & N2Str2Null(txtVisitResults)
        SQL = SQL & " WHERE LogID=" & ENTRY_LOGID
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "LOG PROSPECT VISIT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
    End If


    If ENTRY_LOGID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Visit Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Visit Sucessfully Updated", 500, 1
    End If
    UpdateLog
    rsRefresh
    If ENTRY_LOGID > 0 Then
        RS.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSEARCH
    cmdCancel.Value = True

    If FormExist("MainForm") Then
        MainForm.ShowStatus PROSPECTID
    End If






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
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOG PROSPECT VISIT)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "LOG PROSPECT VISIT")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    InitData
    rsRefresh
    StoreMemVars
    SetEntityDetails PROSPECTID, CustomerCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PROSPECTID = 0
    ENTRY_LOGID = 0
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub LISTVIEW1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH
End Sub

