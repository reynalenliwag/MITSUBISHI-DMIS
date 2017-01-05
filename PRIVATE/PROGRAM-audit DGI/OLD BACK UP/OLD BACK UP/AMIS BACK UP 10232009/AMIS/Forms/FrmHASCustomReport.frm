VERSION 5.00
Begin VB.Form FrmHASCustomReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customized Report(Under Development)"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
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
      Left            =   4080
      MouseIcon       =   "FrmHASCustomReport.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmHASCustomReport.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit Window"
      Top             =   4050
      Width           =   735
   End
   Begin VB.OptionButton OptserviceJouanal 
      Caption         =   "Service Journal Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   4665
   End
   Begin VB.OptionButton OptPartACC 
      Caption         =   "Part and Accessories Journal Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3210
      Width           =   4665
   End
   Begin VB.OptionButton OptPurchases 
      Caption         =   "Summary List Purchases"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2820
      Width           =   4665
   End
   Begin VB.OptionButton OpTransactionRegister 
      Caption         =   "Voucher Payable - Transaction Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Width           =   4665
   End
   Begin VB.OptionButton OptSummary 
      Caption         =   "Trial Balance(Summary)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   4665
   End
   Begin VB.OptionButton OptTB 
      Caption         =   "Trial Balance (Detail Per Voucher)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1650
      Width           =   4665
   End
   Begin VB.OptionButton optOutstanding 
      Caption         =   "Issued and Outstanding - Voucher Payable"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   4665
   End
   Begin VB.OptionButton OptdepositRegister 
      Caption         =   "Cash Deposit Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   870
      Width           =   4665
   End
   Begin VB.OptionButton OptSchedofCheck 
      Caption         =   "Schedule of check issued"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   4665
   End
   Begin VB.OptionButton otpWorkshet 
      Caption         =   "Work Sheet"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "Information: Double click to show report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   30
      TabIndex        =   11
      Top             =   4680
      Width           =   4695
   End
End
Attribute VB_Name = "FrmHASCustomReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
End Sub

Private Sub otpWorkshet_DblClick()
MsgBox "Under development", vbInformation
frmAMISWorkSheet.Show

End Sub
