VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmCSMSMatCheckPrevBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Check Previous Balance, Receipts, Issuance"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   FillColor       =   &H00DEDFDE&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MatCheckPrevBal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5265
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
            Name            =   "Verdana"
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
      Picture         =   "MatCheckPrevBal.frx":01CA
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "MatCheckPrevBal.frx":01E6
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
      MouseIcon       =   "MatCheckPrevBal.frx":0202
      MousePointer    =   99  'Custom
      Picture         =   "MatCheckPrevBal.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exit Window"
      Top             =   1980
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
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
      Height          =   795
      Left            =   3720
      MouseIcon       =   "MatCheckPrevBal.frx":06BA
      MousePointer    =   99  'Custom
      Picture         =   "MatCheckPrevBal.frx":080C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Process Checking of Materials Previous Balance, Receipts and Issuance"
      Top             =   1980
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Issuance of Materials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Total Receipts of Materials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Caption         =   "This Month Moving Ave. Cost"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Last Month Moving Ave. Cost"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Materials This Month On-Hand"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Materials Last Month On-Hand"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
Attribute VB_Name = "frmCSMSMatCheckPrevBal"
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
    If Function_Access(LOGID, "Acess_Process", "CHECK PREVIOUS BALANCE") = False Then Exit Sub

    Dim vLMOH, vOH, vLMMAC                             As Double
    Dim vMAC, vTR, vTI                                 As Double
    Dim i                                              As Integer

    Dim rsMatMas                                       As ADODB.Recordset
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select cost,lastm_mac,onhand,lastm_oh,trecqty,tissqty,mac2 from CSMS_MatMas order by matcde asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        vLMOH = 0: vOH = 0: vLMMAC = 0: vMAC = 0: vTR = 0: vTI = 0
        i = 0
        Do While Not rsMatMas.EOF
            vLMOH = vLMOH + N2Str2Zero(rsMatMas!lastm_oh)
            vOH = vOH + N2Str2Zero(rsMatMas!ONHAND)
            vLMMAC = vLMMAC + (N2Str2Zero(rsMatMas!lastm_mac) * N2Str2Zero(rsMatMas!lastm_oh))
            vMAC = vMAC + (N2Str2Zero(rsMatMas!COST) * N2Str2Zero(rsMatMas!ONHAND))
            vTR = vTR + N2Str2Zero(rsMatMas!trecqty)
            vTI = vTI + N2Str2Zero(rsMatMas!TISSQTY)
            DoEvents
            txtLMOH.Text = Format(vLMOH, DIGIT_FORMAT)
            txtOH.Text = Format(vOH, DIGIT_FORMAT)
            txtLMMAC.Text = Format(vLMMAC, MAXIMUM_DIGIT)
            txtMAC.Text = Format(vMAC, MAXIMUM_DIGIT)
            txtTR.Text = Format(vTR, DIGIT_FORMAT)
            txtTI.Text = Format(vTI, DIGIT_FORMAT)
            i = i + 1
            progCPB.Value = (i / rsMatMas.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsMatMas.MoveNext
        Loop
    Else
        MsgSpeechBox "Error opening Materials Master File"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSMatCheckPrevBal = Nothing
    UnloadForm Me
End Sub
