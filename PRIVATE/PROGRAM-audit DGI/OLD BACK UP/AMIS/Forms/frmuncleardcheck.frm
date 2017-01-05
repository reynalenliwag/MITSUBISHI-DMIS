VERSION 5.00
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frm_TOOLS_UnclearedChecks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank uncleard check"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9450
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   29
      Top             =   6480
      Width           =   1980
      Begin VB.CommandButton cmdCancel 
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
         Left            =   765
         MouseIcon       =   "frmuncleardcheck.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   15
         MouseIcon       =   "frmuncleardcheck.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   1710
      ScaleHeight     =   900
      ScaleWidth      =   12195
      TabIndex        =   16
      Top             =   5610
      Width           =   12195
      Begin VB.CommandButton cmdExit 
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
         Left            =   8505
         MouseIcon       =   "frmuncleardcheck.frx":0932
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":0A84
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   7755
         MouseIcon       =   "frmuncleardcheck.frx":0DEA
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print this Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelCO 
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7005
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmuncleardcheck.frx":12A2
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":13F4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel this Transaction"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6255
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmuncleardcheck.frx":172E
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":1880
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Unpost this Transaction"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5500
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmuncleardcheck.frx":1BC5
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":1D17
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Post this Transaction"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdEdit 
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
         Left            =   4755
         MouseIcon       =   "frmuncleardcheck.frx":203C
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":218E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   4005
         MouseIcon       =   "frmuncleardcheck.frx":24EA
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":263C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         Left            =   3255
         MouseIcon       =   "frmuncleardcheck.frx":294F
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":2AA1
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Move to Last Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   2505
         MouseIcon       =   "frmuncleardcheck.frx":2DF1
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":2F43
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to First Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   1755
         MouseIcon       =   "frmuncleardcheck.frx":32A1
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":33F3
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Find a Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   1005
         MouseIcon       =   "frmuncleardcheck.frx":36ED
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":383F
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move to Next Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   255
         MouseIcon       =   "frmuncleardcheck.frx":3B97
         MousePointer    =   99  'Custom
         Picture         =   "frmuncleardcheck.frx":3CE9
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "List of checks"
      Height          =   3195
      Left            =   90
      TabIndex        =   15
      Top             =   2370
      Width           =   10935
      Begin FlexCell.Grid Grid1 
         Height          =   2835
         Left            =   90
         TabIndex        =   32
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5001
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   30
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bank Info"
      Height          =   2265
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   10935
      Begin VB.TextBox txtend 
         Height          =   315
         Left            =   8910
         TabIndex        =   14
         Top             =   600
         Width           =   1905
      End
      Begin VB.TextBox txtstart 
         Height          =   315
         Left            =   8910
         TabIndex        =   12
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtdescription 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   5625
      End
      Begin VB.TextBox txtaccountcode 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1410
         Width           =   3015
      End
      Begin VB.TextBox txtbankacctno 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1020
         Width           =   3015
      End
      Begin VB.ComboBox cbobank 
         Height          =   330
         Left            =   1800
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   660
         Width           =   5565
      End
      Begin VB.TextBox txtbankcode 
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   3015
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ending:"
         Height          =   255
         Left            =   7380
         TabIndex        =   13
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Start:"
         Height          =   255
         Left            =   7410
         TabIndex        =   11
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Account code:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank Account No:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank Code:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_TOOLS_UnclearedChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbobank_Click()
    displaybankinfo (cboBank.Text)

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    displayBankname
End Sub
Sub displayBankname()
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select bankname from all_banks order by id")
    If Not (RS.EOF And RS.BOF) Then
        cboBank.Clear
        Do While Not RS.EOF
            cboBank.AddItem RS!BankName
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub
Sub displaybankinfo(BankName As String)
    Dim RSinfo                                    As New ADODB.Recordset
    Dim rsdesc                                    As New ADODB.Recordset
    Set RSinfo = gconDMIS.Execute("Select bankcode,bankacctno,acctcode from all_banks where bankname = '" & BankName & "'")
    If Not (RSinfo.EOF And RSinfo.BOF) Then
        txtBankCode.Text = Null2String(RSinfo!bankcode)
        txtbankacctno.Text = Null2String(RSinfo!BankAcctNo)
        txtaccountcode.Text = Null2String(RSinfo!ACCTCODE)
        Set rsdesc = gconDMIS.Execute("SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = '" & txtaccountcode.Text & "'")
        If Not (rsdesc.EOF And rsdesc.BOF) Then
            txtDescription = Null2String(rsdesc!Description)
        Else
            txtDescription = ""
        End If
    End If
    Set RSinfo = Nothing
End Sub

Sub rsRefresh()
'    Dim rs As New ADODB.Recordset
'    set rs

End Sub

