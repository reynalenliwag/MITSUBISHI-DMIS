VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubTitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Sub-Title"
   ClientHeight    =   5070
   ClientLeft      =   1530
   ClientTop       =   1170
   ClientWidth     =   7665
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SubTitle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7665
   Begin VB.PictureBox Picture3 
      Height          =   4980
      Left            =   30
      ScaleHeight     =   4920
      ScaleWidth      =   1845
      TabIndex        =   18
      Top             =   60
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6030
         Left            =   -390
         Picture         =   "SubTitle.frx":0442
         Top             =   -150
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1980
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   13
      Top             =   4140
      Width           =   5625
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4860
         Picture         =   "SubTitle.frx":19AF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4170
         Picture         =   "SubTitle.frx":1DF1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3480
         Picture         =   "SubTitle.frx":2233
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2790
         Picture         =   "SubTitle.frx":2675
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2100
         Picture         =   "SubTitle.frx":2AB7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1410
         Picture         =   "SubTitle.frx":2EF9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   750
         Picture         =   "SubTitle.frx":333B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   675
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         Picture         =   "SubTitle.frx":377D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   1980
      TabIndex        =   12
      Top             =   -30
      Width           =   5625
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   570
         Width           =   4305
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1230
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "XXXX"
         Top             =   180
         Width           =   765
      End
      Begin Crystal.CrystalReport rptHeader 
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Account Headers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -60
         TabIndex        =   21
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -60
         TabIndex        =   17
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3870
         TabIndex        =   16
         Top             =   540
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4350
         TabIndex        =   15
         Top             =   570
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   1980
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   14
      Top             =   4140
      Width           =   5625
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
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
         Height          =   795
         Left            =   4860
         Picture         =   "SubTitle.frx":3BBF
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4170
         Picture         =   "SubTitle.frx":4001
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   1980
      TabIndex        =   19
      Top             =   900
      Width           =   5625
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   22
         Top             =   150
         Width           =   5445
      End
      Begin MSComctlLib.ListView lstSubHeader 
         Height          =   2625
         Left            =   60
         TabIndex        =   20
         Top             =   540
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "SubTitle.frx":4443
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ACCOUNT TYPE"
            Object.Width           =   7761
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSubTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSubHeader As ADODB.Recordset
Dim AddorEdit, PrevCode As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
On Error Resume Next
txtCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: txtCode.Enabled = True: StoreMemvars
End Sub

Private Sub cmdDelete_Click()
If ShowConfirmDelete = True Then
   gconAMIS.Execute "delete * from SubHeader where code = " & N2Str2Null((lstSubHeader.SelectedItem))
   rsRefresh
   StoreMemvars
End If
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True: txtCode.Enabled = False
StoreEntry (lstSubHeader.SelectedItem)
PrevCode = txtCode.Text
On Error Resume Next
txtCode.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
rsSubHeader.MoveNext
If rsSubHeader.EOF Then
   rsSubHeader.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsSubHeader.MovePrevious
If rsSubHeader.BOF Then
   rsSubHeader.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
PrintSQLReport rptSubHeader, AMIS_REPORT_PATH & "SubHeader.rpt", "", AMIS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim VtxtCode, VtxtDescription, VcboType As String
VtxtCode = N2Str2Null(txtCode.Text)
VtxtDescription = N2Str2Null(txtDescription.Text)
If AddorEdit = "ADD" Then
   gconAMIS.Execute "Insert into SubHeader " & _
                    "(code,Description) " & _
                    " values (" & VtxtCode & "," & VtxtDescription & ")"
Else
   gconAMIS.Execute "Update SubHeader set" & _
                    " code = " & VtxtCode & "," & _
                    " Description = " & VtxtDescription & _
                    " where code = '" & PrevCode & "'"
End If
rsRefresh
On Error Resume Next
rsSubHeader.Find "code = " & VtxtCode
cmdCancel.Value = True
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
initMemvars
rsRefresh
StoreMemvars
FillGrid
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsSubHeader = New ADODB.Recordset
Set rsSubHeader = gconAMIS.Execute("select code,description from SubHeader order by code asc")
End Sub

Sub initMemvars()
Frame1.Enabled = True
txtCode.Text = ""
txtDescription.Text = ""
End Sub

Sub StoreMemvars()
If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then
   Frame1.Enabled = False
   txtCode.Text = Null2String(rsSubHeader!code)
   txtDescription.Text = Null2String(rsSubHeader!Description)
Else
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub

Sub StoreEntry(XXX As Variant)
Dim rsSubHeader2 As ADODB.Recordset
Set rsSubHeader2 = New ADODB.Recordset
Set rsSubHeader2 = gconAMIS.Execute("select * from SubHeader where code = '" & XXX & "'")
If Not rsSubHeader2.EOF And Not rsSubHeader2.BOF Then
   txtCode.Text = Null2String(rsSubHeader2!code)
   txtDescription.Text = Null2String(rsSubHeader2!Description)
End If
End Sub

Sub FillGrid()
Dim rsSubHeader2 As ADODB.Recordset
lstSubHeader.Sorted = False: lstSubHeader.ListItems.Clear
Set rsSubHeader2 = New ADODB.Recordset
Set rsSubHeader2 = gconAMIS.Execute("select code,description from SubHeader")
If Not (rsSubHeader2.EOF And rsSubHeader2.BOF) Then
   Listview_Loadval Me.lstSubHeader.ListItems, rsSubHeader2
   lstSubHeader.Refresh
End If
End Sub

Sub FillSearchGrid(XXX As Variant)
Dim rsSubHeader2 As ADODB.Recordset
lstSubHeader.Sorted = False: lstSubHeader.ListItems.Clear
Set rsSubHeader2 = New ADODB.Recordset
Set rsSubHeader2 = gconAMIS.Execute("select code,description from SubHeader where Description like '" & XXX & "%'")
If Not (rsSubHeader2.EOF And rsSubHeader2.BOF) Then
   Listview_Loadval Me.lstSubHeader.ListItems, rsSubHeader2
   lstSubHeader.Refresh
End If
End Sub

Private Sub lstSubHeader_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstSubHeader
     .Sorted = True
     If .SortKey = ColumnHeader.Index - 1 Then
        If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
     Else
        .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
     End If
End With
End Sub

Private Sub lstSubHeader_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstSubHeader_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsSubHeader.Bookmark = rsFind(rsSubHeader.Clone, "code", Str(lstSubHeader.SelectedItem)).Bookmark
StoreMemvars
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
   FillGrid
Else
   FillSearchGrid (txtSearch.Text)
End If
End Sub
