VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmOSMSProcessCheckPrevBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplies Check Previous Balance, Receipts, Issuance"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000F&
   Icon            =   "CheckPrevBal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   5265
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   3270
      ScaleHeight     =   1905
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   90
      Width           =   2235
      Begin VB.TextBox txtTI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   930
         Width           =   1815
      End
      Begin VB.TextBox txtLMMAC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtOH 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtLMOH 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   0
         Width           =   1815
      End
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   2340
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   767
      Picture         =   "CheckPrevBal.frx":030A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "CheckPrevBal.frx":0326
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
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
      Left            =   4440
      MouseIcon       =   "CheckPrevBal.frx":0342
      MousePointer    =   99  'Custom
      Picture         =   "CheckPrevBal.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
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
      Left            =   3720
      MouseIcon       =   "CheckPrevBal.frx":07FA
      MousePointer    =   99  'Custom
      Picture         =   "CheckPrevBal.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Issuances of Supplies"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3165
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipts of Supplies"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   3165
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "This Month Inventory Cost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1050
      Width           =   3165
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month Inventory Cost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   750
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplies This Month On-Hand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplies Last Month On-Hand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3165
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   675
   End
End
Attribute VB_Name = "frmOSMSProcessCheckPrevBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    txtLMOH.Text = 0: txtOH.Text = 0
    txtLMMAC.Text = 0: txtMAC.Text = 0
    txtTR.Text = 0: txtTI.Text = 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdCheck_Click()
    Dim vLMOH, vOH, vLMcost As Double
    Dim vCOST, vTR, vTI As Double
    Dim i As Integer

    Dim rsSupply As ADODB.Recordset
    Set rsSupply = New ADODB.Recordset
    rsSupply.Open "Select cost,lastm_cost,onhand,lastm_oh,trecqty,tissqty from OSMS_SUPPLY order by supply_code asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        vLMOH = 0: vOH = 0: vLMcost = 0: vCOST = 0: vTR = 0: vTI = 0
        i = 0
        Do While Not rsSupply.EOF
            vLMOH = vLMOH + N2Str2Zero(rsSupply!lastm_oh)
            vOH = vOH + N2Str2Zero(rsSupply!Onhand)
            vLMcost = vLMcost + (N2Str2Zero(rsSupply!Lastm_Cost) * N2Str2Zero(rsSupply!lastm_oh))
            vCOST = vCOST + (N2Str2Zero(rsSupply!Cost) * N2Str2Zero(rsSupply!Onhand))
            vTR = vTR + N2Str2Zero(rsSupply!trecqty)
            vTI = vTI + N2Str2Zero(rsSupply!tissqty)
            DoEvents
            txtLMOH.Text = Format(vLMOH, DIGIT_FORMAT)
            txtOH.Text = Format(vOH, DIGIT_FORMAT)
            txtLMMAC.Text = Format(vLMcost, MAXIMUM_DIGIT)
            txtMAC.Text = Format(vCOST, MAXIMUM_DIGIT)
            txtTR.Text = Format(vTR, DIGIT_FORMAT)
            txtTI.Text = Format(vTI, DIGIT_FORMAT)
            i = i + 1
            progCPB.Value = (i / rsSupply.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsSupply.MoveNext
        Loop
    Else
        MsgSpeechBox "Error opening Supplies Master File"
    End If
End Sub

