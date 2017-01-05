VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILECustomer 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master List"
   ClientHeight    =   6225
   ClientLeft      =   585
   ClientTop       =   330
   ClientWidth     =   6825
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AMISCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   6825
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   6150
      Left            =   -2670
      ScaleHeight     =   6090
      ScaleWidth      =   2505
      TabIndex        =   33
      Top             =   60
      Width           =   2565
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   0
         Picture         =   "AMISCustomer.frx":08CA
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   30
      TabIndex        =   22
      Top             =   -30
      Width           =   6765
      Begin VB.TextBox txtCustCode 
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
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtMiddleName 
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
         Left            =   4530
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   930
         Width           =   2115
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   2310
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   930
         Width           =   2115
      End
      Begin VB.TextBox txtCellNo 
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
         Left            =   4080
         MaxLength       =   17
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2400
         Width           =   2565
      End
      Begin VB.TextBox txtPhone 
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
         Left            =   750
         MaxLength       =   17
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2400
         Width           =   2565
      End
      Begin RichTextLib.RichTextBox txtAddress 
         Height          =   525
         Left            =   960
         TabIndex        =   7
         Top             =   1770
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   926
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"AMISCustomer.frx":15242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtLastName 
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
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   930
         Width           =   2115
      End
      Begin VB.TextBox txtAccountNo 
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
         Left            =   4500
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   2145
      End
      Begin VB.TextBox txtCustName 
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
         Left            =   1710
         MaxLength       =   150
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1350
         Width           =   4935
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   6750
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Left            =   0
         TabIndex        =   34
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Left            =   4530
         TabIndex        =   32
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Left            =   2310
         TabIndex        =   31
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cell #"
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
         Left            =   3390
         TabIndex        =   30
         Top             =   2460
         Width           =   645
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6720
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
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
         TabIndex        =   29
         Top             =   2460
         Width           =   645
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
         Left            =   90
         TabIndex        =   28
         Top             =   1770
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No"
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
         Left            =   2970
         TabIndex        =   27
         Top             =   210
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
         TabIndex        =   26
         Top             =   930
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
         TabIndex        =   25
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Left            =   90
         TabIndex        =   24
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   23
         Top             =   1410
         Width           =   1575
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   30
      TabIndex        =   35
      Top             =   2790
      Width           =   6765
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
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   150
         Width           =   6585
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   1875
         Left            =   30
         TabIndex        =   36
         Top             =   540
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3307
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
         MouseIcon       =   "AMISCustomer.frx":152C8
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CUSTOMER NAME"
            Object.Width           =   8290
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   6735
      TabIndex        =   20
      Top             =   5310
      Width           =   6765
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
         Left            =   5850
         MouseIcon       =   "AMISCustomer.frx":1542A
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":1557C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   855
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
         Left            =   5010
         MouseIcon       =   "AMISCustomer.frx":159BE
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":15B10
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   855
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
         Left            =   4170
         MouseIcon       =   "AMISCustomer.frx":15F52
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":160A4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   855
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
         Left            =   3330
         MouseIcon       =   "AMISCustomer.frx":164E6
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":16638
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   855
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
         Left            =   2490
         MouseIcon       =   "AMISCustomer.frx":16A7A
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":16BCC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   855
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
         Left            =   1650
         MouseIcon       =   "AMISCustomer.frx":1700E
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":17160
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   855
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
         Left            =   840
         MouseIcon       =   "AMISCustomer.frx":175A2
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":176F4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   825
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
         Left            =   30
         MouseIcon       =   "AMISCustomer.frx":17B36
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":17C88
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   6735
      TabIndex        =   21
      Top             =   5310
      Width           =   6765
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
         Left            =   5850
         MouseIcon       =   "AMISCustomer.frx":180CA
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":1821C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   30
         Width           =   855
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
         Left            =   5010
         MouseIcon       =   "AMISCustomer.frx":1865E
         MousePointer    =   99  'Custom
         Picture         =   "AMISCustomer.frx":187B0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILECustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer, rsCusCtl As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
Dim NewCtlCde As String
Dim rsCustomer As ADODB.Recordset
Dim k As Integer
Screen.MousePointer = 11
gconCSMIOS.Execute "delete from cusctl"
For k = 65 To 90
    Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select custcode from customer where left(custcode,1) = '" & Chr(k) & "' order by custcode desc", gconCSMIOS
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
       NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!custcode, 2, 5)) + 1, "00000")
       gconCSMIOS.Execute "insert into cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
    Else
       gconCSMIOS.Execute "insert into cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
    End If
Next
Screen.MousePointer = 0
initMemvars
lstCustomer.Enabled = False
Picture1.Visible = False
Picture2.Visible = True
On Error Resume Next
txtLastName.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
lstCustomer.Enabled = True
fraDetails.Enabled = True
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
   gconCSMIOS.Execute "delete from Customer where id = " & labid.Caption
End If
rsRefresh
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
fraDetails.Enabled = False
lstCustomer.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
txtSearch.SetFocus
'Dim findStr As String
'findStr = InputBox("Please Input Customer ...", "Find")
'If findStr <> "" Then
'   On Error Resume Next
'   rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "CustName", findStr).Bookmark
'   If Err.Number = 3021 Then
'      On Error Resume Next
'      rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "LastName", findStr).Bookmark
'      If Err.Number = 3021 Then
'         On Error GoTo ErrorCustCode
'         rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "FirstName", findStr).Bookmark
'      End If
'   End If
'End If
'StoreMemvars
'Exit Sub

'ErrorCustCode:
'If Err.Number = 3021 Then
'   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
'   Resume Next
'End If
End Sub

Private Sub cmdNext_Click()
rsCustomer.MoveNext
If rsCustomer.EOF Then
   rsCustomer.MoveLast
   MsgBox "Last of Record"
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsCustomer.MovePrevious
If rsCustomer.BOF Then
   rsCustomer.MoveFirst
   MsgBox "Beginning of record"
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
'PrintReport rptCustomer, AMIS_REPORT_PATH & "Customer.rpt", "", 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCustCode
Dim VtxtCustCode, VtxtAccountNo, VtxtCustName, VTXTLastName As String
Dim VTXTFirstName, VtxtMiddleName, VTXTAddress As String
Dim VtxtPhone, VtxtCellNo, NewCtlCde As String

VtxtCustCode = N2Str2Null(txtCustCode.Text)
VtxtAccountNo = N2Str2Null(txtAccountNo.Text)
VtxtCustName = N2Str2Null(txtCustName.Text)
VTXTLastName = N2Str2Null(txtLastName.Text)
VTXTFirstName = N2Str2Null(txtFirstName.Text)
VtxtMiddleName = N2Str2Null(txtMiddleName.Text)
VTXTAddress = N2Str2Null(txtAddress.Text)
VtxtPhone = N2Str2Null(txtPhone.Text)
VtxtCellNo = N2Str2Null(txtCellNo.Text)
NewCtlCde = Left(txtCustCode.Text, 1) & Format(NumericVal(Mid(txtCustCode.Text, 2, 5)) + 1, "00000")

If AddorEdit = "ADD" Then
   Dim rsCustomerDup As ADODB.Recordset
   Set rsCustomerDup = New ADODB.Recordset
       rsCustomerDup.Open "select AccountNo from Customer where AccountNo = " & VtxtAccountNo, gconCSMIOS
   If Not rsCustomerDup.EOF And Not rsCustomerDup.BOF Then
      MsgBox "Customer Account No. Already Exist!", vbCritical, "Duplicate Account No. Not Allowed"
      Exit Sub
   End If
   Set rsCustomerDup = New ADODB.Recordset
       rsCustomerDup.Open "select AccountNo from Customer where CustCode = " & VtxtCustCode, gconCSMIOS
   If Not rsCustomerDup.EOF And Not rsCustomerDup.BOF Then
      MsgBox "Customer Code Already Exist!", vbCritical, "Duplicate Customer Code Not Allowed"
      Exit Sub
   End If
   gconCSMIOS.Execute "Insert into Customer " & _
                    "(CustCode,AccountNo,CustName,Lastname,Firstname,middlename,address,phone,CellNo) " & _
                    " values (" & VtxtCustCode & "," & VtxtAccountNo & _
                    ", " & VtxtCustName & ", " & VTXTLastName & _
                    ", " & VTXTFirstName & ", " & VtxtMiddleName & ", " & VTXTAddress & _
                    ", " & VtxtPhone & ", " & VtxtCellNo & ")"
Else
   gconCSMIOS.Execute "Update Customer set" & _
                    " CustCode = " & VtxtCustCode & "," & _
                    " AccountNo = " & VtxtAccountNo & "," & _
                    " CustName = " & VtxtCustName & "," & _
                    " Lastname = " & VTXTLastName & "," & _
                    " Firstname = " & VTXTFirstName & "," & _
                    " Middlename = " & VtxtMiddleName & "," & _
                    " address = " & VTXTAddress & "," & _
                    " phone = " & VtxtPhone & "," & _
                    " CellNo = " & VtxtCellNo & _
                    " where ID = " & labid.Caption
End If
gconCSMIOS.Execute "update cusctl set ctlcde = '" & NewCtlCde & "' where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'"
rsRefresh
On Error Resume Next
rsCustomer.Find "CustCode = " & VtxtCustCode
cmdCancel.Value = True
Exit Sub

ErrorCustCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCustCode As Integer, Shift As Integer)
MoveKeyPress KeyCustCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
'DrawXPCtl Me
rsRefresh
txtSearch.Text = ""
initMemvars
StoreMemvars
FillGrid
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from Customer WHERE CUSTCODE <> '999999' order by CUSTCODE asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
FillGrid
End Sub

Sub initMemvars()
Frame1.Enabled = True
'Dim rsCustomerAcc As ADODB.Recordset
'Set rsCustomerAcc = New ADODB.Recordset
'    rsCustomerAcc.Open "select custcode from customer order by custcode asc", gconCSMIOS
'If Not rsCustomerAcc.EOF And Not rsCustomerAcc.BOF Then
'   rsCustomerAcc.MoveLast
'   txtCustCode.Text = Format(N2Str2Zero(rsCustomerAcc!CUSTCODE) + 1, "0000")
'Else
'   txtCustCode.Text = "0001"
'End If
txtCustCode.Text = ""
txtAccountNo.Text = ""
txtCustName.Text = ""
txtLastName.Text = ""
txtFirstName.Text = ""
txtMiddleName.Text = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtCellNo.Text = ""
End Sub

Sub StoreMemvars()
If Not rsCustomer.EOF And Not rsCustomer.BOF Then
   Frame1.Enabled = False
   labid.Caption = rsCustomer!ID
   txtCustCode.Text = Null2String(rsCustomer!custcode)
   txtAccountNo.Text = Null2String(rsCustomer!ACCOUNTNO)
   txtCustName.Text = Null2String(rsCustomer!custname)
   txtLastName.Text = Null2String(rsCustomer!lastname)
   txtFirstName.Text = Null2String(rsCustomer!firstname)
   txtMiddleName.Text = Null2String(rsCustomer!middlename)
   txtAddress.Text = Null2String(rsCustomer!Address)
   txtPhone.Text = Null2String(rsCustomer!PHONE)
   txtCellNo.Text = Null2String(rsCustomer!CELLNO)
Else
   MsgBox "No Such Record!"
   cmdAdd.Value = True
End If
End Sub

Sub FillGrid()
Dim rsCustomer2 As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomer2 = New ADODB.Recordset
Set rsCustomer2 = gconCSMIOS.Execute("select CustCode,CustName from Customer where custcode <> '999999' ORDER BY CUSTCODE ASC")
If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
   lstCustomer.Refresh
   lstCustomer.Enabled = True
Else
   lstCustomer.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsCustomer2 As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomer2 = New ADODB.Recordset
Set rsCustomer2 = gconCSMIOS.Execute("select CustCode,CustName from Customer where custcode <> '999999' and CustName like '" & XXX & "%' ORDER BY CUSTCODE ASC")
If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
   lstCustomer.Refresh
   lstCustomer.Enabled = True
Else
   lstCustomer.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If PMIOS_ORDER_SHOW = True Then
   frmPMIOSCustomerOrder.txtCustCode.Text = txtCustCode.Text
   frmPMIOSCustomerOrder.txtCustName.Text = txtCustName.Text & vbCrLf & txtAddress.Text
End If
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

Private Sub lstCustomer_GotFocus()
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "custcode", lstCustomer.SelectedItem).Bookmark
StoreMemvars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "custcode", lstCustomer.SelectedItem).Bookmark
StoreMemvars
End Sub

Private Sub txtCustName_Change()
'If Len(txtCustName.Text) = 1 Then
'   If AddorEdit = "ADD" Then
'      Set rsCusCtl = New ADODB.Recordset
'      Set rsCusCtl = gconCSMIOS.Execute("select ctlcde from cusctl where left(ctlcde,1) = '" & Left(txtCustName.Text, 1) & "'")
'      If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then txtCustCode.Text = Null2String(rsCusCtl!ctlcde)
'   End If
'End If
If Len(txtCustName.Text) = 1 Then
   If AddorEdit = "ADD" Then
      Set rsCusCtl = New ADODB.Recordset
      Set rsCusCtl = gconCSMIOS.Execute("select ctlcde from cusctl where left(ctlcde,1) = '" & Left(txtCustName.Text, 1) & "'")
      If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then txtCustCode.Text = Null2String(rsCusCtl!ctlcde)
   End If
End If
End Sub

Private Sub txtCustName_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtFirstName_Change()
txtCustName.Text = Trim(UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleName.Text))
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtLastName_Change()
If Len(txtLastName.Text) = 1 Then
   If AddorEdit = "ADD" Then
      Set rsCusCtl = New ADODB.Recordset
      Set rsCusCtl = gconCSMIOS.Execute("select ctlcde from cusctl where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'")
      If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then txtCustCode.Text = Null2String(rsCusCtl!ctlcde)
   End If
End If
txtCustName.Text = Trim(UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleName.Text))
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtMiddleName_Change()
txtCustName.Text = Trim(UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleName.Text))
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
If Trim(txtSearch.Text) = "" Then
   FillGrid
Else
   FillSearchGrid (txtSearch.Text)
End If
End Sub

Private Sub txtSearch_KeyDown(KeyCustCode As Integer, Shift As Integer)
If KeyCustCode = vbKeyDown Then lstCustomer.SetFocus
End Sub

Private Sub txtAddress_LostFocus()
'txtAddress.Text = Cap1st(txtAddress.Text)
End Sub

Private Sub txtCustName_LostFocus()
txtCustName.Text = UCase(txtCustName.Text)
End Sub

Private Sub txtFirstName_LostFocus()
'txtFirstName.Text = Cap1st(txtFirstName.Text)
End Sub

Private Sub txtLastName_LostFocus()
'txtLastName.Text = Cap1st(txtLastName.Text)
End Sub

