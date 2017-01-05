VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILETaxRate 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Rate Master List"
   ClientHeight    =   4200
   ClientLeft      =   1125
   ClientTop       =   705
   ClientWidth     =   7980
   ForeColor       =   &H00DEDFDE&
   Icon            =   "TaxRate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   7980
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   60
      ScaleHeight     =   3990
      ScaleWidth      =   2175
      TabIndex        =   21
      Top             =   90
      Width           =   2235
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   0
         Picture         =   "TaxRate.frx":08CA
         Top             =   0
         Width           =   2550
      End
   End
   Begin Crystal.CrystalReport rptBanks 
      Left            =   7440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Banks"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2310
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   12
      Top             =   3240
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
         MouseIcon       =   "TaxRate.frx":15242
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":15394
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
         MouseIcon       =   "TaxRate.frx":157D6
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":15928
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
         MouseIcon       =   "TaxRate.frx":15D6A
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":15EBC
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
         MouseIcon       =   "TaxRate.frx":162FE
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":16450
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
         MouseIcon       =   "TaxRate.frx":16892
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":169E4
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
         MouseIcon       =   "TaxRate.frx":16E26
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":16F78
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
         MouseIcon       =   "TaxRate.frx":173BA
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":1750C
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
         MouseIcon       =   "TaxRate.frx":1794E
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":17AA0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2310
      TabIndex        =   15
      Top             =   0
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   540
         Width           =   3945
      End
      Begin VB.TextBox txtTaxRateCode 
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
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   2070
         TabIndex        =   22
         Top             =   210
         Width           =   315
      End
      Begin VB.Label Label5 
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
         Left            =   -30
         TabIndex        =   19
         Top             =   600
         Width           =   1605
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Rate Code"
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
         TabIndex        =   16
         Top             =   210
         Width           =   1725
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   2310
      TabIndex        =   14
      Top             =   870
      Width           =   5625
      Begin MSComctlLib.ListView lstTaxRate 
         Height          =   2115
         Left            =   30
         TabIndex        =   20
         Top             =   150
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3731
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
         MouseIcon       =   "TaxRate.frx":17EE2
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TAX RATE CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
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
      Left            =   2310
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   13
      Top             =   3240
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
         MouseIcon       =   "TaxRate.frx":18044
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":18196
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
         MouseIcon       =   "TaxRate.frx":185D8
         MousePointer    =   99  'Custom
         Picture         =   "TaxRate.frx":1872A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILETaxRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTaxRate As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
initMemvars
Picture1.Visible = False
Picture2.Visible = True
lstTaxRate.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
lstTaxRate.Enabled = True
fraDetails.Enabled = True
StoreMemvars
FillGrid
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
   gconAmis.Execute "delete from TaxRate where ID = " & lstTaxRate.SelectedItem.SubItems(2)
End If
rsRefresh
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
StoreEntry (lstTaxRate.SelectedItem.SubItems(2))
lstTaxRate.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim findStr As String
findStr = InputBox("Please Input TaxRate ...", "Find")
If findStr <> "" Then
   On Error Resume Next
   rsTaxRate.Bookmark = rsFind(rsTaxRate.Clone, "TaxRateCode", findStr).Bookmark
   If Err.Number = 3021 Then
      On Error GoTo ErrorCode
      rsTaxRate.Bookmark = rsFind(rsTaxRate.Clone, "Description", findStr).Bookmark
   End If
End If
StoreMemvars
Exit Sub

ErrorCode:
If Err.Number = 3021 Then
   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
   Resume Next
End If
End Sub

Private Sub cmdNext_Click()
rsTaxRate.MoveNext
If rsTaxRate.EOF Then
   rsTaxRate.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsTaxRate.MovePrevious
If rsTaxRate.BOF Then
   rsTaxRate.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
'PrintSQLReport rptTaxRate, AMIS_REPORT_PATH & "TaxRate.rpt", "", AMIS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim VtxtTaxRateCode, VtxtDescription As String

VtxtTaxRateCode = N2Str2Null(txtTaxRateCode.Text)
VtxtDescription = N2Str2Null(txtDescription.Text)

If AddorEdit = "ADD" Then
   Dim rsTaxRateDup As ADODB.Recordset
   Set rsTaxRateDup = New ADODB.Recordset
       rsTaxRateDup.Open "select TaxRateCode from TaxRate where TaxRateCode = " & VtxtTaxRateCode, gconAmis
   If Not rsTaxRateDup.EOF And Not rsTaxRateDup.BOF Then
      MsgBox "Bank Code Already Exist!", vbCritical, "Duplicate Bank Code Not Allowed"
      Exit Sub
   End If
   gconAmis.Execute "Insert into TaxRate " & _
                    "(TaxRateCode,Description) " & _
                    " values (" & VtxtTaxRateCode & _
                    ", " & VtxtDescription & ")"
Else
   gconAmis.Execute "Update TaxRate set" & _
                    " TaxRateCode = " & VtxtTaxRateCode & ", " & _
                    " Description = " & VtxtDescription & _
                    " where ID = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsTaxRate.Find "ID = " & labID.Caption
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
FillGrid
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsTaxRate = New ADODB.Recordset
    rsTaxRate.Open "select * from TaxRate order by TaxRateCode asc", gconAmis, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
Frame1.Enabled = True
txtTaxRateCode.Text = ""
txtDescription.Text = ""
End Sub

Sub StoreMemvars()
If Not rsTaxRate.EOF And Not rsTaxRate.BOF Then
   Frame1.Enabled = False
   labID.Caption = rsTaxRate!ID
   txtTaxRateCode.Text = Null2String(rsTaxRate!TaxRateCode)
   txtDescription.Text = Null2String(rsTaxRate!Description)
Else
   lstTaxRate.ListItems.Clear
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub

Sub StoreEntry(XXX As Variant)
Dim rsTaxRate2 As ADODB.Recordset
Set rsTaxRate2 = New ADODB.Recordset
    rsTaxRate2.Open "select * from TaxRate where ID = " & XXX, gconAmis, adOpenForwardOnly, adLockReadOnly
If Not rsTaxRate2.EOF And Not rsTaxRate2.BOF Then
   fraDetails.Enabled = False
   lstTaxRate.Enabled = False
   labID.Caption = rsTaxRate2!ID
   txtTaxRateCode.Text = Null2String(rsTaxRate2!TaxRateCode)
   txtDescription.Text = Null2String(rsTaxRate2!Description)
End If
End Sub

Sub FillGrid()
Dim rsTaxRate2 As ADODB.Recordset
lstTaxRate.Sorted = False: lstTaxRate.ListItems.Clear
Set rsTaxRate2 = New ADODB.Recordset
Set rsTaxRate2 = gconAmis.Execute("select TaxRateCode,Description,ID from TaxRate")
If Not (rsTaxRate2.EOF And rsTaxRate2.BOF) Then
   lstTaxRate.Enabled = True
   Listview_Loadval Me.lstTaxRate.ListItems, rsTaxRate2
   lstTaxRate.Refresh
Else
   lstTaxRate.Enabled = False
End If
End Sub

Private Sub lstTaxRate_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstTaxRate
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

Private Sub lstTaxRate_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstTaxRate_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsTaxRate.Bookmark = rsFind(rsTaxRate.Clone, "DESCRIPTION", lstTaxRate.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub txtTaxRateCode_LostFocus()
txtTaxRateCode.Text = UCase(txtTaxRateCode.Text)
End Sub
