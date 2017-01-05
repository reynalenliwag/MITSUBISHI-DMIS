VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payee Master List"
   ClientHeight    =   4830
   ClientLeft      =   585
   ClientTop       =   540
   ClientWidth     =   7875
   ForeColor       =   &H8000000F&
   Icon            =   "Payee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7875
   Begin VB.PictureBox Picture3 
      Height          =   4725
      Left            =   30
      ScaleHeight     =   4665
      ScaleWidth      =   1845
      TabIndex        =   24
      Top             =   60
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6030
         Left            =   -390
         Picture         =   "Payee.frx":0442
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
      TabIndex        =   15
      Top             =   3900
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
         Left            =   4830
         Picture         =   "Payee.frx":19AF
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   4140
         Picture         =   "Payee.frx":1DF1
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   3450
         Picture         =   "Payee.frx":2233
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   2760
         Picture         =   "Payee.frx":2675
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   2070
         Picture         =   "Payee.frx":2AB7
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   1380
         Picture         =   "Payee.frx":2EF9
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   720
         Picture         =   "Payee.frx":333B
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   60
         Picture         =   "Payee.frx":377D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   1980
      TabIndex        =   18
      Top             =   -30
      Width           =   5865
      Begin VB.TextBox txtAddress 
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
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   4425
      End
      Begin VB.TextBox txtPayeeCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtPayeeName 
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
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   570
         Width           =   4425
      End
      Begin Crystal.CrystalReport rptPayee 
         Left            =   2160
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Payee Master List"
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   510
         TabIndex        =   23
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payee Code"
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
         Left            =   -150
         TabIndex        =   22
         Top             =   240
         Width           =   1455
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
         Left            =   4710
         TabIndex        =   21
         Top             =   600
         Width           =   225
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
         Left            =   3780
         TabIndex        =   20
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payee Name"
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
         Left            =   -270
         TabIndex        =   19
         Top             =   630
         Width           =   1575
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   1980
      TabIndex        =   17
      Top             =   1290
      Width           =   5865
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   150
         Width           =   5715
      End
      Begin MSComctlLib.ListView lstPayee 
         Height          =   1995
         Left            =   30
         TabIndex        =   4
         Top             =   540
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PAYEE CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PAYEE NAME"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
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
      TabIndex        =   16
      Top             =   3900
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
         Left            =   4830
         Picture         =   "Payee.frx":3BBF
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   4140
         Picture         =   "Payee.frx":4001
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmPayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPayee As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
initMemvars
Picture1.Visible = False
Picture2.Visible = True
lstPayee.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
lstPayee.Enabled = True
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
   gconAMIS.Execute "delete * from Payee where id = " & lstPayee.SelectedItem.SubItems(2)
End If
rsRefresh
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
StoreEntry (lstPayee.SelectedItem.SubItems(2))
lstPayee.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
rsPayee.MoveNext
If rsPayee.EOF Then
   rsPayee.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsPayee.MovePrevious
If rsPayee.BOF Then
   rsPayee.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
'PrintReport rptPayee, AMIS_REPORT_PATH & "Payee.rpt", "", 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim VtxtPayeeCode, VtxtPayeeName, VtxtAddress As String

VtxtPayeeCode = N2Str2Null(txtPayeeCode.Text)
VtxtPayeeName = N2Str2Null(txtPayeeName.Text)
VtxtAddress = N2Str2Null(txtAddress.Text)

If AddorEdit = "ADD" Then
   Dim rsPayeeDup As ADODB.Recordset
   Set rsPayeeDup = New ADODB.Recordset
       rsPayeeDup.Open "select PayeeCode from Payee where PayeeCode = " & VtxtPayeeCode, gconAMIS
   If Not rsPayeeDup.EOF And Not rsPayeeDup.BOF Then
      MsgBox "Payee Code Already Exist!", vbCritical, "Duplicate Payee Code Not Allowed"
      Exit Sub
   End If
   gconAMIS.Execute "Insert into Payee " & _
                    "(PayeeCode,PayeeName,address,Profile_ID) " & _
                    " values (" & VtxtPayeeCode & _
                    ", " & VtxtPayeeName & ", " & VtxtAddress & ",1)"
Else
   gconAMIS.Execute "Update Payee set" & _
                    " PayeeName = " & VtxtPayeeName & "," & _
                    " address = " & VtxtAddress & _
                    " where PayeeCode = " & VtxtPayeeCode
End If
rsRefresh
On Error Resume Next
rsPayee.Find "PayeeCode = " & VtxtPayeeCode
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
rsRefresh
initMemvars
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsPayee = New ADODB.Recordset
    rsPayee.Open "select * from Payee order by PayeeCode asc", gconAMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
Frame1.Enabled = True
Dim rsPayeeAcc As ADODB.Recordset
Set rsPayeeAcc = New ADODB.Recordset
    rsPayeeAcc.Open "select PayeeCode from Payee order by PayeeCode asc", gconAMIS
If Not rsPayeeAcc.EOF And Not rsPayeeAcc.BOF Then
   rsPayeeAcc.MoveLast
   txtPayeeCode.Text = Format(N2Str2Zero(rsPayeeAcc!PayeeCode) + 1, "0000")
Else
   txtPayeeCode.Text = "0001"
End If
txtPayeeName.Text = ""
txtAddress.Text = ""
txtSearch.Text = ""
End Sub

Sub StoreMemvars()
If Not rsPayee.EOF And Not rsPayee.BOF Then
   Frame1.Enabled = False
   labid.Caption = rsPayee!ID
   txtPayeeCode.Text = Null2String(rsPayee!PayeeCode)
   txtPayeeName.Text = Null2String(rsPayee!PayeeName)
   txtAddress.Text = Null2String(rsPayee!address)
   FillGrid
Else
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub

Sub StoreEntry(XXX As Variant)
Dim rsPayee2 As ADODB.Recordset
Set rsPayee2 = New ADODB.Recordset
    rsPayee2.Open "select * from Payee where ID = " & XXX, gconAMIS, adOpenForwardOnly, adLockReadOnly
If Not rsPayee2.EOF And Not rsPayee2.BOF Then
   labid.Caption = rsPayee2!ID
   txtPayeeCode.Text = Null2String(rsPayee2!PayeeCode)
   txtPayeeName.Text = Null2String(rsPayee2!PayeeName)
   txtAddress.Text = Null2String(rsPayee2!address)
End If
End Sub

Sub FillGrid()
Dim rsPayee2 As ADODB.Recordset
lstPayee.Sorted = False: lstPayee.ListItems.Clear
Set rsPayee2 = New ADODB.Recordset
Set rsPayee2 = gconAMIS.Execute("select payeecode,payeename,ID from payee")
If Not (rsPayee2.EOF And rsPayee2.BOF) Then
   Listview_Loadval Me.lstPayee.ListItems, rsPayee2
   lstPayee.Refresh
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsPayee2 As ADODB.Recordset
lstPayee.Sorted = False: lstPayee.ListItems.Clear
Set rsPayee2 = New ADODB.Recordset
Set rsPayee2 = gconAMIS.Execute("select payeecode,payeename,ID from payee where payeename like '" & XXX & "%'")
If Not (rsPayee2.EOF And rsPayee2.BOF) Then
   Listview_Loadval Me.lstPayee.ListItems, rsPayee2
   lstPayee.Refresh
End If
End Sub

Private Sub lstPayee_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstPayee
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

Private Sub lstPayee_DblClick()
cmdEdit.Value = True
End Sub

Private Sub txtAddress_LostFocus()
txtAddress.Text = Cap1st(txtAddress.Text)
End Sub

Private Sub txtPayeeName_LostFocus()
txtPayeeName.Text = UCase(txtPayeeName.Text)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
   FillGrid
Else
   FillSearchGrid (txtSearch.Text)
End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstPayee.SetFocus
End Sub
