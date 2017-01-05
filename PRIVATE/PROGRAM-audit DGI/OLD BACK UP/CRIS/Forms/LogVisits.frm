VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Log_Visits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Visits"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
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
   ScaleHeight     =   7095
   ScaleWidth      =   7650
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7650
      TabIndex        =   23
      Top             =   6180
      Width           =   7650
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5970
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   24
         Top             =   0
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   755
            MouseIcon       =   "LogVisits.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "LogVisits.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2190
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   27
         Top             =   0
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "LogVisits.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "LogVisits.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "LogVisits.frx":1B31
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":1C83
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "LogVisits.frx":1FDF
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2131
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "LogVisits.frx":2444
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2596
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1080
            MouseIcon       =   "LogVisits.frx":2890
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":29E2
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   390
            MouseIcon       =   "LogVisits.frx":2D3A
            MousePointer    =   99  'Custom
            Picture         =   "LogVisits.frx":2E8C
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7650
      TabIndex        =   10
      Top             =   0
      Width           =   7650
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Text            =   "LogVisits.frx":31EB
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   2955
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
         TabIndex        =   5
         Top             =   1035
         Width           =   4275
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
         TabIndex        =   4
         Top             =   2580
         Width           =   4425
      End
      Begin MSComCtl2.DTPicker txtVisitDate 
         Height          =   345
         Left            =   150
         TabIndex        =   6
         Top             =   465
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20250625
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   9
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   8
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   2280
         Width           =   645
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1785
      Width           =   2955
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   195
         Width           =   2880
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3765
         Left            =   45
         TabIndex        =   2
         Top             =   645
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   6641
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
   End
End
Attribute VB_Name = "frmCRIS_Log_Visits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                                         As Long
Dim ENTRY_LOGID                                        As Long
Dim rs                                                 As ADODB.Recordset
Dim CustomerCode                                       As String

Friend Sub AddVisit(xProsID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    ProspectID = xProsID
    CustomerCode = xCustCode
End Sub

Private Sub cmdAdd_Click()
    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    On Error Resume Next
    txtVisitDate.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    picSearch.Enabled = True
    ENTRY_LOGID = 0
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Prospect_Visits where Logid=" & ENTRY_LOGID
        UpdateLog
        FillSearchGrid txtSearch
        rsRefresh
        StoreMemvars


    End If
End Sub

Private Sub cmdEdit_Click()
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    On Error Resume Next
    txtVisitDate.SetFocus

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdSave_Click()
    Dim SQL                                            As String
    If ENTRY_LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospect_Visits (ProspectID, CSCDE, DateTimeVisit, Comments,Results) VALUES ("
        SQL = SQL & ProspectID & ", "
        SQL = SQL & N2Str2Null(CustomerCode) & ", "
        SQL = SQL & N2Str2Null(txtVisitDate.Value) & ", "
        SQL = SQL & N2Str2Null(txtVisitComments) & ", "
        SQL = SQL & N2Str2Null(txtVisitResults) & ")"

    Else
        SQL = "Update CRIS_Prospect_Visits SET  "
        SQL = SQL & " DateTimeVisit=" & N2Str2Null(txtVisitDate.Value) & " , "
        SQL = SQL & " Comments=" & N2Str2Null(txtVisitComments) & " , "
        SQL = SQL & " Results=" & N2Str2Null(txtVisitResults)
        SQL = SQL & " WHERE LogID=" & ENTRY_LOGID
    End If

    gconDMIS.Execute (SQL)
    If ENTRY_LOGID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Visit Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Visit Sucessfully Updated", 500, 1
    End If
    UpdateLog
    rsRefresh
    If ENTRY_LOGID > 0 Then
        rs.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSearch
    cmdCancel.Value = True



End Sub

Sub FillSearchGrid(xxx As String)
    Dim TEMPRS                                         As ADODB.Recordset
    If CustomerCode <> vbNullString Then
        Set TEMPRS = gconDMIS.Execute("SELECT DATETIMEVISIT,COMMENTS , LOGID  FROM CRIS_PROSPECT_VISITS " & _
                                    " WHERE  CSCDE=" & N2Str2Null(CustomerCode) & " AND  CONVERT(VARCHAR, DATETIMEVISIT , 101)  LIKE  '" & ReplaceQuote(xxx) & "%' ORDER BY DATETIMEVISIT  ASC")
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT DATETIMEVISIT,COMMENTS , LOGID  FROM CRIS_PROSPECT_VISITS " & _
                                    " WHERE  PROSPECTID=" & ProspectID & " AND  CONVERT(VARCHAR, DATETIMEVISIT , 101)  LIKE  '" & ReplaceQuote(xxx) & "%' ORDER BY DATETIMEVISIT  ASC")
    End If


    flex_FillListView TEMPRS, ListView1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitMemVars
    InitData
    rsRefresh
    StoreMemvars
    SetEntityDetails ProspectID, CustomerCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    ENTRY_LOGID = 0
End Sub

Sub InitData()

    picDataEntry.Enabled = False
    picSearch.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False

    AddColumnHeader "Date , EmailAddress", ListView1
    ResizeColumnHeader ListView1, "40,55"
    FillSearchGrid ""


End Sub

Sub InitMemVars()
    txtVisitDate = DateValue(Now)
    txtVisitComments = ""
    txtVisitResults = ""


End Sub

Private Sub LISTVIEW1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
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

Private Sub LISTVIEW1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rs.MoveFirst
    rs.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemvars
End Sub





Sub rsRefresh()
    Set rs = New ADODB.Recordset
    If CustomerCode <> vbNullString Then
        rs.Open "SELECT * From CRIS_Prospect_Visits Where CSCDE=" & N2Str2Null(CustomerCode) & "Order BY DateTimeVisit desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rs.Open "SELECT * From CRIS_Prospect_Visits Where ProspectID=" & ProspectID & "Order BY DateTimeVisit desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TEMPRS                                         As ADODB.Recordset
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

Sub StoreMemvars()
    If Not rs.EOF And Not rs.BOF Then
        'SELECT ENTRY_LOGID, ProspectID, DateEmail, EmailFrom, EmailTO, Subject, Body, Bound FROM DMIS.dbo.CRIS_Prospect_Email
        ENTRY_LOGID = rs!LOGID
        ProspectID = rs!ProspectID
        txtVisitComments = Null2String(rs!Comments)
        If IsNull(rs!DateTimeVisit) = False Then
            txtVisitDate = rs!DateTimeVisit
        End If
        txtVisitResults = Null2String(rs!Results)

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid txtSearch
End Sub

Sub UpdateLog()
    Dim TSQL                                           As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DateTimeVisit) FROM CRIS_Prospect_VISITS  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGVISIT=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGVISIT=NULL  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

