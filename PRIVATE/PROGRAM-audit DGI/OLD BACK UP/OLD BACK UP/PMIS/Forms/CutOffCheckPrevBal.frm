VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMIS_Physical_CutOffCheckPrevBal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Cut-Off Previous Balance"
   ClientHeight    =   2250
   ClientLeft      =   450
   ClientTop       =   435
   ClientWidth     =   5790
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CutOffCheckPrevBal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5790
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
      Left            =   4920
      MouseIcon       =   "CutOffCheckPrevBal.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CutOffCheckPrevBal.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit Window"
      Top             =   1380
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
      Left            =   4200
      MouseIcon       =   "CutOffCheckPrevBal.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "CutOffCheckPrevBal.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Process Checking of Previous Cut-Off Balance"
      Top             =   1380
      Width           =   735
   End
   Begin VB.PictureBox PicCHKPrev 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   645
      Left            =   3840
      ScaleHeight     =   645
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   90
      Width           =   2235
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   330
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   30
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   5
      Top             =   720
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
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   8
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   9
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
            MICON           =   "CutOffCheckPrevBal.frx":0BAF
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "CutOffCheckPrevBal.frx":0BCB
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "CutOffCheckPrevBal.frx":0BE7
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
         TabIndex        =   11
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Cut-Off Total Moving Ave. Cost"
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
      Top             =   450
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Cut-Off Total On-Hand"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPMIS_Physical_CutOffCheckPrevBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CheckBalance()
    Dim vOH                                                           As Double
    Dim vMAC                                                          As Double
    Dim I                                                             As Integer
    Dim RCOUNT                                                        As Long
    Dim rsCUTOFF                                                      As ADODB.Recordset
    Set rsCUTOFF = New ADODB.Recordset
    rsCUTOFF.Open "Select mac,onhand,PARTNO from CUTOFF  order by PARTNO asc", gconINVENTORY, adOpenKeyset, adLockReadOnly
    If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
        rsCUTOFF.MoveFirst
        vOH = 0: vMAC = 0:
        I = 0
        RCOUNT = rsCUTOFF.RecordCount
        Do While Not rsCUTOFF.EOF
            labProcessing.Caption = "Processing " & DESC_TYPE & " Number: " & Null2String(rsCUTOFF!PARTNO)
            DoEvents
            vOH = vOH + N2Str2IntZero(rsCUTOFF!ONHAND)
            vMAC = vMAC + (N2Str2Zero(rsCUTOFF!Mac) * N2Str2IntZero(rsCUTOFF!ONHAND))
            DoEvents
            txtOH.Text = Format(vOH, DIGIT_FORMAT)
            txtMAC.Text = Format(vMAC, MAXIMUM_DIGIT)
            I = I + 1
            progCPB.Value = (I / RCOUNT) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsCUTOFF.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
    Else
        MsgSpeechBox "Error Opening Part Master File"
    End If
    Set rsCUTOFF = Nothing
End Sub

Private Sub cmdCheck_Click()


    CheckBalance
    LogAudit "R", "CUT OFF PREV BAL"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtOH.Text = 0: txtMAC.Text = 0
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISCheckPrevBal = Nothing
    UnloadForm Me
End Sub

