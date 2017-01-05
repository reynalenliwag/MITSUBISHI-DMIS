VERSION 5.00
Begin VB.Form frmIncomes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incomes"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ForeColor       =   &H8000000F&
   Icon            =   "Income.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2010
   ScaleWidth      =   5700
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   14
      Top             =   1080
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
         Picture         =   "Income.frx":0442
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
         Picture         =   "Income.frx":0884
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
         Picture         =   "Income.frx":0CC6
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
         Picture         =   "Income.frx":1108
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
         Picture         =   "Income.frx":154A
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
         Picture         =   "Income.frx":198C
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
         Picture         =   "Income.frx":1DCE
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
         Picture         =   "Income.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   30
      TabIndex        =   12
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
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   570
         Width           =   3975
      End
      Begin VB.TextBox txtIncomeCode 
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   495
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
         Left            =   30
         TabIndex        =   18
         Top             =   630
         Width           =   1455
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
         TabIndex        =   17
         Top             =   600
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
         TabIndex        =   16
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Income Code"
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
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   15
      Top             =   1080
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
         Picture         =   "Income.frx":2652
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
         Picture         =   "Income.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmIncomes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsIncomes As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
initMemvars
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
txtIncomeCode.Enabled = True
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
   gconAMIS.Execute "delete * from Incomes where IncomeCode = " & N2Str2Null(txtIncomeCode.Text)
End If
rsRefresh
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
txtIncomeCode.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim findStr As String
findStr = InputBox("Please Input Incomes ...", "Find")
If findStr <> "" Then
   On Error GoTo ErrorIncomeCode
   rsIncomes.Bookmark = rsFind(rsIncomes.Clone, "Description", findStr).Bookmark
End If
StoreMemvars
Exit Sub

ErrorIncomeCode:
If Err.Number = 3021 Then
   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
   Resume Next
End If
End Sub

Private Sub cmdNext_Click()
rsIncomes.MoveNext
If rsIncomes.EOF Then
   rsIncomes.MoveLast
   MsgBox "Last of Record"
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsIncomes.MovePrevious
If rsIncomes.BOF Then
   rsIncomes.MoveFirst
   MsgBox "Beginning of record"
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
'PrintReport rptIncomes, AMIS_REPORT_PATH & "Incomes.rpt", "", 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorIncomeCode
Dim VtxtIncomeCode, VtxtDescription As String

VtxtIncomeCode = N2Str2Null(txtIncomeCode.Text)
VtxtDescription = N2Str2Null(txtDescription.Text)
If AddorEdit = "ADD" Then
   Dim rsIncomesDup As ADODB.Recordset
   Set rsIncomesDup = New ADODB.Recordset
       rsIncomesDup.Open "select IncomeCode from Incomes where IncomeCode = " & VtxtIncomeCode, gconAMIS
   If Not rsIncomesDup.EOF And Not rsIncomesDup.BOF Then
      MsgBox "Account IncomeCode Already Exist!", vbCritical, "Duplicate IncomeCode Not Allowed"
      Exit Sub
   End If
   gconAMIS.Execute "Insert into Incomes " & _
                    "(IncomeCode,Description,Profile_ID) " & _
                    " values (" & VtxtIncomeCode & ", " & VtxtDescription & ",1)"
Else
   gconAMIS.Execute "Update Incomes set" & _
                    " Description = " & VtxtDescription & _
                    " where IncomeCode = " & VtxtIncomeCode
End If
rsRefresh
On Error Resume Next
rsIncomes.Find "IncomeCode = " & VtxtIncomeCode
cmdCancel.Value = True
Exit Sub

ErrorIncomeCode:
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
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsIncomes = New ADODB.Recordset
    rsIncomes.Open "select * from Incomes order by IncomeCode", gconAMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
Frame1.Enabled = True
txtIncomeCode.Text = ""
txtDescription.Text = ""
End Sub

Sub StoreMemvars()
If Not rsIncomes.EOF And Not rsIncomes.BOF Then
   Frame1.Enabled = False
   txtIncomeCode.Text = Null2String(rsIncomes!IncomeCode)
   txtDescription.Text = Null2String(rsIncomes!Description)
Else
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub
