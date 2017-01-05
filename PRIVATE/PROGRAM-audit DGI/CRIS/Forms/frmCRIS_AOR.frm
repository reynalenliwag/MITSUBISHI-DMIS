VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCRIS_AOR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AOR Calculator"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3060
      MouseIcon       =   "frmCRIS_AOR.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmCRIS_AOR.frx":0152
      TabIndex        =   14
      Top             =   2790
      Width           =   765
   End
   Begin VB.ComboBox cboTerm 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   180
      ScaleHeight     =   1305
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   1380
      Width           =   4455
      Begin VB.TextBox txtAOR 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   1950
         TabIndex        =   2
         Text            =   "0"
         Top             =   0
         Width           =   825
      End
      Begin MSMask.MaskEdBox txtNetMoAmort 
         Height          =   345
         Left            =   1950
         TabIndex        =   3
         Top             =   870
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFinBaltoFinanced 
         Height          =   345
         Left            =   1950
         TabIndex        =   4
         Top             =   480
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bal. to be financed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   510
         Width           =   1965
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Mo. Amort."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   900
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      MouseIcon       =   "frmCRIS_AOR.frx":0490
      MousePointer    =   99  'Custom
      Picture         =   "frmCRIS_AOR.frx":05E2
      TabIndex        =   0
      Top             =   2790
      Width           =   765
   End
   Begin MSMask.MaskEdBox txtDownPayment 
      Height          =   345
      Left            =   2130
      TabIndex        =   9
      Top             =   570
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   609
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777215
      ForeColor       =   7347754
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtNetSalesPrice 
      Height          =   345
      Left            =   2115
      TabIndex        =   10
      Top             =   180
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   609
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   16777215
      ForeColor       =   7347754
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TERM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   1245
      TabIndex        =   13
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label29 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Down Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   465
      TabIndex        =   12
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label Label30 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Sales Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   465
      TabIndex        =   11
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "frmCRIS_AOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event LineAOR(NetSalesPrice, DownPayment, Term, AOR, FinBaltoFinanced, NetMoAmort)

Private Sub cboTerm_Change()
    Dim NoOfMonths       As Integer
    Dim AOR              As Double
    NoOfMonths = NumericVal(Left(cboTerm.Text, 2))
    If NoOfMonths = 12 Then AOR = 7.61
    If NoOfMonths = 18 Then AOR = 10.48
    If NoOfMonths = 24 Then AOR = 17.45
    If NoOfMonths = 36 Then AOR = 25.55
    If NoOfMonths = 48 Then AOR = 33.96
    If NoOfMonths = 60 Then AOR = 44.15
    txtAOR.Text = AOR
    txtFinBaltoFinanced.Text = ToDoubleNumber(NumericVal(txtNetSalesPrice.Text) - NumericVal(txtDownPayment.Text))
    txtNetMoAmort.Text = ToDoubleNumber(NumericVal(txtFinBaltoFinanced.Text) * (1 + (AOR / 100)) / NoOfMonths)
End Sub

Private Sub cboTerm_Click()
    Dim NoOfMonths       As Integer
    Dim AOR              As Double
    NoOfMonths = CInt(Left(cboTerm.Text, 2))
    If NoOfMonths = 12 Then AOR = 7.61
    If NoOfMonths = 18 Then AOR = 10.48
    If NoOfMonths = 24 Then AOR = 17.45
    If NoOfMonths = 36 Then AOR = 25.55
    If NoOfMonths = 48 Then AOR = 33.96
    If NoOfMonths = 60 Then AOR = 44.15
    txtAOR.Text = AOR
    txtFinBaltoFinanced.Text = CDbl(CDbl(NumericVal(txtNetSalesPrice.Text)) - CDbl(NumericVal(txtDownPayment.Text)))
    txtNetMoAmort.Text = CDbl(CDbl(txtFinBaltoFinanced.Text) * (1 + (AOR / 100)) / NoOfMonths)
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub Command1_Click()
    RaiseEvent LineAOR(txtNetSalesPrice, txtDownPayment, cboTerm, txtAOR, txtFinBaltoFinanced, txtNetMoAmort)
    
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    cboTerm.Clear
    cboTerm.AddItem "12 mos."
    cboTerm.AddItem "18 mos."
    cboTerm.AddItem "24 mos."
    cboTerm.AddItem "36 mos."
    cboTerm.AddItem "48 mos."
    cboTerm.AddItem "60 mos."
End Sub

