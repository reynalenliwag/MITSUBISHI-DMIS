VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISARCheckingTool 
   Caption         =   "AR Checking Tool"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvDetails 
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.OptionButton optAP 
      Caption         =   "AP Accounts"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton optAR 
      Caption         =   "AR Accounts"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.ComboBox cboDescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   2415
      TabIndex        =   3
      Top             =   1680
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   20578307
      CurrentDate     =   38148
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   4980
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   20578307
      CurrentDate     =   38148
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SL Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "GL Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblGL 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      TabIndex        =   5
      Top             =   1710
      Width           =   405
   End
   Begin VB.Label lblAccountCode 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   1635
   End
End
Attribute VB_Name = "frmAMISARCheckingTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsChartAccount                                     As ADODB.Recordset

Private Sub cboDescription_Click()
    lblAccountCode = AccountCode
End Sub

Private Sub Form_Load()
    InitChart

End Sub

Sub InitChart()
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Do While Not rsChartAccount.EOF
            cboDescription.AddItem Null2String(rsChartAccount!DESCRIPTION)
            rsChartAccount.MoveNext
        Loop
    End If
    Set rsChartAccount = Nothing
End Sub

Function AccountCode() As String
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND DESCRIPTION='" & cboDescription.Text & "' ORDER BY ACCTCODE", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function
