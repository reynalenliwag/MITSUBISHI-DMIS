VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmPMIOSRECONCheckPrevBal 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECONCILE Check Previous Balance, Receipts, Issuance"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "RECONCheckPrevBal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "RECONCheckPrevBal.frx":030A
   ScaleHeight     =   4500
   ScaleWidth      =   5835
   Begin VB.PictureBox PicCHKPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   3870
      Picture         =   "RECONCheckPrevBal.frx":3046
      ScaleHeight     =   2865
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   90
      Width           =   2235
      Begin VB.TextBox txtLastY_OH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtTotalRR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1890
         Width           =   1815
      End
      Begin VB.TextBox txtTotalISS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2190
         Width           =   1815
      End
      Begin VB.TextBox txtTI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtTR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   930
         Width           =   1815
      End
      Begin VB.TextBox txtLMMAC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtOH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtLMOH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   795
      Left            =   4800
      MouseIcon       =   "RECONCheckPrevBal.frx":5D82
      MousePointer    =   99  'Custom
      Picture         =   "RECONCheckPrevBal.frx":608C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Check"
      Height          =   795
      Left            =   3900
      MouseIcon       =   "RECONCheckPrevBal.frx":6396
      MousePointer    =   99  'Custom
      Picture         =   "RECONCheckPrevBal.frx":66A0
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Check"
      Top             =   3660
      Width           =   915
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   60
      Picture         =   "RECONCheckPrevBal.frx":6F6A
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   15
      Top             =   3000
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   16
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   17
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         Picture         =   "RECONCheckPrevBal.frx":9CA6
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   18
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   19
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "RECONCheckPrevBal.frx":C9E2
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "RECONCheckPrevBal.frx":C9FE
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "RECONCheckPrevBal.frx":CA1A
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
      Begin VB.Label labCPB 
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
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   21
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Year On-Hand"
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
      Left            =   90
      TabIndex        =   26
      Top             =   2640
      Width           =   3165
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "YTD Total Receipts"
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
      Left            =   90
      TabIndex        =   23
      Top             =   1980
      Width           =   3165
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "YTD Total Issuance"
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
      Left            =   90
      TabIndex        =   22
      Top             =   2280
      Width           =   3165
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "MTD Issuance"
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
      Left            =   90
      TabIndex        =   7
      Top             =   1680
      Width           =   3165
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "MTD Receipts"
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
      Left            =   90
      TabIndex        =   6
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
      Left            =   90
      TabIndex        =   5
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
      Left            =   90
      TabIndex        =   4
      Top             =   750
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "This Month On-Hand"
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
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month On-Hand"
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
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmPMIOSRECONCheckPrevBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
CheckBalance
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
txtLMOH.Text = 0:   txtOH.Text = 0
txtLMMAC.Text = 0:  txtMAC.Text = 0
txtTR.Text = 0:     txtTI.Text = 0
txtTotalRR.Text = 0: txtTotalISS.Text = 0
txtLastY_OH.Text = 0
Screen.MousePointer = 0
End Sub

Sub CheckBalance()
Dim vLYOH, vLMOH, vOH, vLMMAC As Double
Dim vMAC, vTR, vTI, vTotalRR, vTotalISS As Double
Dim i As Integer

Dim rsPartmas As ADODB.Recordset
Set rsPartmas = New ADODB.Recordset
Set rsPartmas = gconPMIOS.Execute("Select SUM(lasty_oh) AS TOTAL_LASTY_OH, SUM(lastm_oh) AS TOTAL_LASTM_OH, SUM(lastm_mac * lastm_oh) TOTAL_LASTM_MAC_ONHAND, SUM(onhand) AS TOTAL_ONHAND, SUM(mac * onhand) TOTAL_MAC_ONHAND, SUM(trecqty) AS TOTAL_TRECQTY,SUM(tissqty) AS TOTAL_TISSQTY,SUM(receipts) TOTAL_RECEIPTS,SUM(issuances) AS TOTAL_ISSUANCES from NEW_PARTMAS")
If Not rsPartmas.EOF And Not rsPartmas.BOF Then
   'rsPartMas.MoveFirst
   'vLMOH = 0:   vOH = 0:   vLMMAC = 0:   vMAC = 0:   vTR = 0:   vTI = 0
   'vTotalRR = 0: vTotalISS = 0
   'i = 0
   'Do While Not rsPartMas.EOF
   '   labProcessing.Caption = "Processing Part Number: " & Null2String(rsPartMas!PartNo)
   '   DoEvents
      vLYOH = vLYOH + N2Str2IntZero(rsPartmas!TOTAL_LASTY_OH)
      vLMOH = vLMOH + N2Str2IntZero(rsPartmas!TOTAL_LASTM_OH)
      vOH = vOH + N2Str2IntZero(rsPartmas!TOTAL_ONHAND)
      vLMMAC = N2Str2Zero(rsPartmas!TOTAL_LASTM_MAC_ONHAND)
      vMAC = N2Str2Zero(rsPartmas!TOTAL_MAC_ONHAND)
      vTR = vTR + N2Str2IntZero(rsPartmas!TOTAL_trecqty)
      vTI = vTI + N2Str2IntZero(rsPartmas!TOTAL_tissqty)
      vTotalRR = vTotalRR + N2Str2IntZero(rsPartmas!TOTAL_receipts)
      vTotalISS = vTotalISS + N2Str2IntZero(rsPartmas!TOTAL_issuances)
      DoEvents
      txtLMOH.Text = Format(vLMOH, DIGIT_FORMAT)
      txtOH.Text = Format(vOH, DIGIT_FORMAT)
      txtLMMAC.Text = Format(vLMMAC, MAXIMUM_DIGIT)
      txtMAC.Text = Format(vMAC, MAXIMUM_DIGIT)
      txtTR.Text = Format(vTR, DIGIT_FORMAT)
      txtTI.Text = Format(vTI, DIGIT_FORMAT)
      txtTotalRR.Text = Format(vTotalRR, DIGIT_FORMAT)
      txtTotalISS.Text = Format(vTotalISS, DIGIT_FORMAT)
      txtLastY_OH.Text = Format(vLYOH, DIGIT_FORMAT)
      'i = i + 1
      'progCPB.Value = (i / rsPartMas.RecordCount) * 100
      progCPB.Value = 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      'DoEvents
      'rsPartMas.MoveNext
   'Loop
   labProcessing.Caption = ""
   DoEvents
Else
   MsgSpeechBox "Error Opening Part Master File"
End If
Set rsPartmas = Nothing
End Sub

Sub OLDCheckBalance()
Dim vLMOH, vOH, vLMMAC As Double
Dim vMAC, vTR, vTI, vTotalRR, vTotalISS As Double
Dim i As Integer

Dim rsPartmas As ADODB.Recordset
Set rsPartmas = New ADODB.Recordset
    rsPartmas.Open "Select mac,lastm_mac,onhand,lastm_oh,trecqty,tissqty,mac2,partno,receipts,issuances from NEW_PARTMAS order by partno asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsPartmas.EOF And Not rsPartmas.BOF Then
   rsPartmas.MoveFirst
   vLMOH = 0:   vOH = 0:   vLMMAC = 0:   vMAC = 0:   vTR = 0:   vTI = 0
   vTotalRR = 0: vTotalISS = 0
   i = 0
   Do While Not rsPartmas.EOF
      labProcessing.Caption = "Processing Part Number: " & Null2String(rsPartmas!PartNo)
      DoEvents
      vLMOH = vLMOH + N2Str2IntZero(rsPartmas!lastm_oh)
      vOH = vOH + N2Str2IntZero(rsPartmas!Onhand)
      vLMMAC = vLMMAC + (N2Str2Zero(rsPartmas!lastm_mac) * N2Str2IntZero(rsPartmas!lastm_oh))
      vMAC = vMAC + (N2Str2Zero(rsPartmas!MAC) * N2Str2IntZero(rsPartmas!Onhand))
      vTR = vTR + N2Str2IntZero(rsPartmas!trecqty)
      vTI = vTI + N2Str2IntZero(rsPartmas!tissqty)
      vTotalRR = vTotalRR + N2Str2IntZero(rsPartmas!receipts)
      vTotalISS = vTotalISS + N2Str2IntZero(rsPartmas!issuances)
      DoEvents
      txtLMOH.Text = Format(vLMOH, DIGIT_FORMAT)
      txtOH.Text = Format(vOH, DIGIT_FORMAT)
      txtLMMAC.Text = Format(vLMMAC, MAXIMUM_DIGIT)
      txtMAC.Text = Format(vMAC, MAXIMUM_DIGIT)
      txtTR.Text = Format(vTR, DIGIT_FORMAT)
      txtTI.Text = Format(vTI, DIGIT_FORMAT)
      txtTotalRR.Text = Format(vTotalRR, DIGIT_FORMAT)
      txtTotalISS.Text = Format(vTotalISS, DIGIT_FORMAT)
      i = i + 1
      progCPB.Value = (i / rsPartmas.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsPartmas.MoveNext
   Loop
   labProcessing.Caption = ""
   DoEvents
Else
   MsgSpeechBox "Error Opening Part Master File"
End If
Set rsPartmas = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSCheckPrevBal = Nothing
UnloadForm Me
End Sub
