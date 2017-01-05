VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISCustomer 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Data Entry"
   ClientHeight    =   5505
   ClientLeft      =   1935
   ClientTop       =   435
   ClientWidth     =   9105
   ForeColor       =   &H00F5F5F5&
   Icon            =   "customer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9105
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2700
      ScaleHeight     =   825
      ScaleWidth      =   6315
      TabIndex        =   35
      Top             =   4560
      Width           =   6345
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F2EFE9&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5490
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdSelect 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4710
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3930
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3150
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":18E0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2370
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":21AA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1590
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":2A74
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   810
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":333E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Picture         =   "customer.frx":3780
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   2700
      TabIndex        =   22
      Top             =   60
      Width           =   6345
      Begin VB.TextBox txtCusAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   735
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "customer.frx":3BC2
         Top             =   2280
         Width           =   6075
      End
      Begin VB.ComboBox cboCusNam 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1590
         Width           =   6105
      End
      Begin VB.TextBox txtCuscde 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   750
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtCuscat 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3630
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4020
         Width           =   2595
      End
      Begin VB.TextBox txtCusphon1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1440
         MaxLength       =   17
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4020
         Width           =   1815
      End
      Begin VB.TextBox txtCuszipc 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   180
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4020
         Width           =   855
      End
      Begin VB.TextBox txtProvAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   180
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3330
         Width           =   6045
      End
      Begin VB.TextBox txtMiddleInitial 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5730
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2940
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   930
         Width           =   2700
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   150
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   930
         Width           =   2700
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Full Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   1320
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3660
         TabIndex        =   30
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Provincial Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   3060
         Width           =   2085
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   2010
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5730
         TabIndex        =   25
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2910
         TabIndex        =   24
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   660
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7380
      ScaleHeight     =   825
      ScaleWidth      =   1635
      TabIndex        =   36
      Top             =   4560
      Width           =   1665
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   810
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":3BC8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F2EFE9&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   30
         MaskColor       =   &H0000FFFF&
         Picture         =   "customer.frx":4C0A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   60
      TabIndex        =   37
      Top             =   60
      Width           =   2565
      Begin VB.TextBox textSearch 
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
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   150
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   4755
         Left            =   60
         TabIndex        =   21
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   8387
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   15920873
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
         MouseIcon       =   "customer.frx":504C
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   3000
      TabIndex        =   32
      Top             =   450
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3330
      TabIndex        =   31
      Top             =   450
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCMISCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCusmas, rsRepor, rsCusCtl As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
initMemvars
On Error Resume Next
txtLastName.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
AddorEdit = ""
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsCusmas.BOF Or Not rsCusmas.EOF Then
   If ShowConfirmDelete = True Then
      gconCSMIOS.Execute "delete from Cusmas where id = " & labID.Caption
      ShowDeletedMsg
   End If
Else
   ShowNothingToDeleteMsg
End If
rsRefresh
StoreMemvars
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
On Error Resume Next
txtLastName.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Picture3.Visible = False
textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
rsCusmas.MoveNext
If rsCusmas.EOF Then
   rsCusmas.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsCusmas.MovePrevious
If rsCusmas.BOF Then
   rsCusmas.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode

If IsNull(txtCUSCDE.Text) = True Then
   MsgSpeechBox "Customer Code must not be empty"
   On Error Resume Next
   txtCUSCDE.SetFocus
   Exit Sub
Else
   If AddorEdit = "ADD" Then
      Dim rsfindDup As ADODB.Recordset
      Set rsfindDup = New ADODB.Recordset
          rsfindDup.Open "select cuscde from cusmas where cuscde = '" & txtCUSCDE.Text & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
      If Not rsfindDup.EOF And Not rsfindDup.BOF Then
         MsgSpeechBox "Customer Code already exist!"
         On Error Resume Next
         txtCUSCDE.SetFocus
         Exit Sub
      End If
   End If
End If
Dim NewCtlCde, CustomerNam As String

NewCtlCde = Left(txtCUSCDE.Text, 1) & Format(NumericVal(Mid(txtCUSCDE.Text, 2, 5)) + 1, "00000")
CustomerNam = N2Str2Null(UCase(txtLastName.Text))

Dim VTXTCuscde, VTXTLastName, VTXTFirstname As String
Dim VTXTMiddleInitial, VTXTCusadd, VTXTProvadd As String
Dim VTXTCuszipc, VTXTCusphon1, VTXTCuscat As String
Dim VTXTCusComp, VTXTCusNam As String

VTXTCuscde = N2Str2Null(txtCUSCDE.Text)
VTXTLastName = N2Str2Null(UCase(txtLastName.Text))
VTXTFirstname = N2Str2Null(UCase(txtFirstName.Text))
VTXTMiddleInitial = N2Str2Null(txtMiddleInitial.Text)
VTXTCusComp = N2Str2Null(cboCusNam.Text)
VTXTCusNam = N2Str2Null(cboCusNam.Text)
VTXTCusadd = N2Str2Null(txtCusAdd.Text)
VTXTProvadd = N2Str2Null(txtProvAdd.Text)
VTXTCuszipc = N2Str2Null(txtCuszipc.Text)
VTXTCusphon1 = N2Str2Null(txtCusphon1.Text)
VTXTCuscat = NumericVal(txtCuscat.Text)

If AddorEdit = "ADD" Then
   Dim rsCusMasDup As ADODB.Recordset
   Set rsCusMasDup = New ADODB.Recordset
       rsCusMasDup.Open "select id from cusmas order by id asc", gconCSMIOS
   If Not rsCusMasDup.EOF And Not rsCusMasDup.BOF Then
      rsCusMasDup.MoveLast
      labID.Caption = NumericVal(rsCusMasDup!Id) + 1
   End If
   gconCSMIOS.Execute "Insert into Cusmas" & _
                    " (cuscde,cuscomp,cusnam,Lastname,Firstname,MiddleInitial,Cusadd,Provadd,cuszipc,cusphon1,cuscat,usercode,lastupdate,timeupdate)" & _
                    " values (" & VTXTCuscde & ", " & VTXTCusComp & ", " & VTXTCusNam & ", " & VTXTLastName & ", " & VTXTFirstname & ", " & VTXTMiddleInitial & ", " & _
                    " " & VTXTCusadd & ", " & VTXTProvadd & ", " & VTXTCuszipc & ", " & VTXTCusphon1 & ", " & VTXTCuscat & ", '" & LOGCODE & "', '" & Date & "', '" & Time & "')"
Else
   gconCSMIOS.Execute "update Cusmas set" & _
                    " Cuscde = " & VTXTCuscde & "," & _
                    " cuscomp = " & VTXTCusComp & "," & _
                    " cusnam = " & VTXTCusNam & "," & _
                    " Lastname = " & VTXTLastName & "," & _
                    " Firstname = " & VTXTFirstname & "," & _
                    " MiddleInitial = " & VTXTMiddleInitial & "," & _
                    " Cusadd = " & VTXTCusadd & "," & _
                    " Provadd = " & VTXTProvadd & "," & _
                    " cuszipc = " & VTXTCuszipc & "," & _
                    " cusphon1 = " & VTXTCusphon1 & "," & _
                    " cuscat = " & VTXTCuscat & "," & _
                    " usercode = '" & LOGCODE & "'," & _
                    " lastupdate = '" & LOGDATE & "'," & _
                    " timeupdate = '" & Time & "'" & _
                    " where id = " & labID.Caption
End If
gconCSMIOS.Execute "update cusctl set ctlcde = '" & NewCtlCde & "' where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'"
rsRefresh
On Error Resume Next
rsCusmas.Find "id =" & labID.Caption
cmdCancel.Value = True
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
rsRefresh
Frame1.Enabled = False
initMemvars
textSearch.Text = "": Picture3.ZOrder 0
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub initMemvars()
txtCUSCDE.Text = ""
txtLastName.Text = ""
txtFirstName.Text = ""
txtMiddleInitial.Text = ""
txtCusAdd.Text = ""
txtProvAdd.Text = ""
txtCuszipc.Text = ""
txtCusphon1.Text = ""
txtCuscat.Text = ""
FillCboCusnam
End Sub

Sub StoreMemvars()
If Not rsCusmas.EOF And Not rsCusmas.BOF Then
   labID.Caption = rsCusmas!Id
   txtCUSCDE.Text = Null2String(rsCusmas!Cuscde)
   txtLastName.Text = Null2String(rsCusmas!Lastname)
   txtFirstName.Text = Null2String(rsCusmas!Firstname)
   cboCusNam.Text = Null2String(rsCusmas!CusNam)
   txtMiddleInitial.Text = Null2String(rsCusmas!MiddleInitial)
   txtCusAdd.Text = Null2String(rsCusmas!Cusadd)
   txtProvAdd.Text = Null2String(rsCusmas!Provadd)
   txtCuszipc.Text = Null2String(rsCusmas!cuszipc)
   txtCusphon1.Text = Null2String(rsCusmas!cusphon1)
   txtCuscat.Text = Null2String(rsCusmas!cuscat)
Else
   ShowNoRecord
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsCusmas = New ADODB.Recordset
    rsCusmas.Open "select * from Cusmas order by cusnam asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
CURRENT_CUST_CODE = txtCUSCDE.Text
End Sub

Private Sub txtFirstName_Change()
cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Private Sub txtLastname_Change()
If Len(txtLastName.Text) = 1 Then
   If AddorEdit = "ADD" Then
      Set rsCusCtl = New ADODB.Recordset
         rsCusCtl.Open "select ctlcde from cusctl where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
      If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
          txtCUSCDE.Text = Null2String(rsCusCtl!ctlcde)
      End If
   End If
End If
cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Private Sub txtMiddleInitial_Change()
cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Sub FillCboCusnam()
Dim rsCusMas2 As ADODB.Recordset
Set rsCusMas2 = New ADODB.Recordset
    rsCusMas2.Open "select cusnam from cusmas", gconCSMIOS
If Not rsCusMas2.EOF And Not rsCusMas2.BOF Then
   rsCusMas2.MoveFirst
   cboCusNam.Clear
   Do While Not rsCusMas2.EOF
      cboCusNam.AddItem Null2String(rsCusMas2!CusNam)
      rsCusMas2.MoveNext
   Loop
End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsCusmas.Bookmark = rsFind(rsCusmas.Clone, "cuscde", lstCustomer.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstCustomer
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

Private Sub lstCustomer_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
If Trim(textSearch.Text) = "" Then
   FillGrid
Else
   FillSearchGrid (textSearch.Text)
End If
End Sub

Sub FillGrid()
Dim rsCustomer As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomer = New ADODB.Recordset
Set rsCustomer = gconCSMIOS.Execute("select cusnam,cuscde from cusmas")
If Not (rsCustomer.EOF And rsCustomer.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
   lstCustomer.Refresh
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsCustomer As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomer = New ADODB.Recordset
Set rsCustomer = gconCSMIOS.Execute("select cusnam,cuscde from Cusmas where cusnam like'" & XXX & "%'")
If Not (rsCustomer.EOF And rsCustomer.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
   lstCustomer.Refresh
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstCustomer.SetFocus
End Sub
