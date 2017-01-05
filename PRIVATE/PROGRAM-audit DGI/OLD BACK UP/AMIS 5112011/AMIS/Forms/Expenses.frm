VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expenses"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ForeColor       =   &H8000000F&
   Icon            =   "Expenses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   5700
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5595
      TabIndex        =   15
      Top             =   1440
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
         Picture         =   "Expenses.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "Expenses.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "Expenses.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "Expenses.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "Expenses.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "Expenses.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "Expenses.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "Expenses.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   30
      TabIndex        =   13
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
      Begin VB.TextBox txtExpCode 
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
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   20
         Top             =   1020
         Width           =   975
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Code"
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
         TabIndex        =   14
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
      TabIndex        =   16
      Top             =   1440
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
         Picture         =   "Expenses.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "Expenses.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsExpenses As ADODB.Recordset
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
txtExpCode.Enabled = True
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
   gconAMIS.Execute "delete * from Expenses where ExpCode = " & N2Str2Null(txtExpCode.Text)
End If
rsRefresh
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
txtExpCode.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim findStr As String
findStr = InputBox("Please Input Expenses ...", "Find")
If findStr <> "" Then
   On Error GoTo ErrorExpCode
   rsExpenses.Bookmark = rsFind(rsExpenses.Clone, "Description", findStr).Bookmark
End If
StoreMemvars
Exit Sub

ErrorExpCode:
If Err.Number = 3021 Then
   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
   Resume Next
End If
End Sub

Private Sub cmdNext_Click()
rsExpenses.MoveNext
If rsExpenses.EOF Then
   rsExpenses.MoveLast
   MsgBox "Last of Record"
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsExpenses.MovePrevious
If rsExpenses.BOF Then
   rsExpenses.MoveFirst
   MsgBox "Beginning of record"
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
'PrintReport rptExpenses, AMIS_REPORT_PATH & "Expenses.rpt", "", 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorExpCode
Dim VtxtExpCode, VtxtDescription As String
Dim VtxtAmount As Double
VtxtExpCode = N2Str2Null(txtExpCode.Text)
VtxtDescription = N2Str2Null(txtDescription.Text)
VtxtAmount = N2Str2Zero(txtDescription.Text)
If AddorEdit = "ADD" Then
   Dim rsExpensesDup As ADODB.Recordset
   Set rsExpensesDup = New ADODB.Recordset
       rsExpensesDup.Open "select ExpCode from Expenses where ExpCode = " & VtxtExpCode, gconAMIS
   If Not rsExpensesDup.EOF And Not rsExpensesDup.BOF Then
      MsgBox "Account Expense Code Already Exist!", vbCritical, "Duplicate Expense Code Not Allowed"
      Exit Sub
   End If
   gconAMIS.Execute "Insert into Expenses " & _
                    "(ExpCode,Description,Amount,Profile_ID) " & _
                    " values (" & VtxtExpCode & ", " & VtxtDescription & ", " & VtxtAmount & ",1)"
Else
   gconAMIS.Execute "Update Expenses set" & _
                    " Description = " & VtxtDescription & "," & _
                    " Amount = " & VtxtAmount & _
                    " where ExpCode = " & VtxtExpCode
End If
rsRefresh
On Error Resume Next
rsExpenses.Find "ExpCode = " & VtxtExpCode
cmdCancel.Value = True
Exit Sub

ErrorExpCode:
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
Set rsExpenses = New ADODB.Recordset
    rsExpenses.Open "select * from Expenses order by ExpCode", gconAMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
Frame1.Enabled = True
txtExpCode.Text = ""
txtDescription.Text = ""
txtAmount.Text = 0#
End Sub

Sub StoreMemvars()
If Not rsExpenses.EOF And Not rsExpenses.BOF Then
   Frame1.Enabled = False
   txtExpCode.Text = Null2String(rsExpenses!ExpCode)
   txtDescription.Text = Null2String(rsExpenses!Description)
   txtAmount.Text = Null2String(rsExpenses!amount)
Else
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub
