VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmOSMSProcessUpdateAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Supplies Adjustment File"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000F&
   Icon            =   "UpdateAdjustment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   5805
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
      MouseIcon       =   "UpdateAdjustment.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAdjustment.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
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
      MouseIcon       =   "UpdateAdjustment.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAdjustment.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateAdjustment.frx":0BAF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateAdjustment.frx":0BCB
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   3
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
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
            MICON           =   "UpdateAdjustment.frx":0BE7
         End
      End
      Begin VB.Label labCPB 
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
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmOSMSProcessUpdateAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdjustment As ADODB.Recordset

Private Sub cmdCheck_Click()
    Set rsAdjustment = New ADODB.Recordset
    rsAdjustment.Open "select * from OSMS_Adjustment where status = 'N' order by Supply_Code asc", gconDMIS
    If rsAdjustment.EOF And rsAdjustment.BOF Then
        MsgSpeechBox "Error: Adjustment File is Empty or Adjustments had been Posted already!"
        Exit Sub
    Else
        cmdCheck.Enabled = False
        cmdExit.Enabled = False
        DoEvents
        UpdateAdjustment
        cmdExit.Enabled = True
        DoEvents
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
End Sub

Sub UpdateAdjustment()
    Dim i As Integer
    Dim vSupply_Code As String
    Dim vID As Integer
    Dim vAdd, vMinus As Double
    Dim VStatus As String

    rsAdjustment.MoveFirst
    Screen.MousePointer = 11
    Me.Caption = "Updating Adjustment to Transactions Master File"
    DoEvents
    i = 0
    Do While Not rsAdjustment.EOF
        vID = rsAdjustment!Id
        vSupply_Code = N2Str2Null(rsAdjustment!Supply_Code)
        vMinus = N2Str2Zero(rsAdjustment!minus)
        vAdd = N2Str2Zero(rsAdjustment!Add)
        VStatus = "'N'"
        If vAdd <> 0 Then
            gconDMIS.Execute "Insert into OSMS_rrDetails " & _
                             "(Supply_Code,status,rrQuantity,RRNumber,item_no)" & _
                           " values (" & vSupply_Code & ", 'N'," & vAdd & ",'111111','1111')"
        Else
            gconDMIS.Execute "Insert into ISSUANCE_DETAILS " & _
                             "(Supply_Code,status,ID_QUANTITY,TRANS_NO,ID_ITEM_NO)" & _
                           " values (" & vSupply_Code & ", 'N'," & vMinus & ",'000000','0000')"
        End If
        gconDMIS.Execute "UPDATE OSMS_ADJUSTMENT set status = 'P' where id = " & vID
        DoEvents
        i = i + 1
        progCPB.Value = (i / rsAdjustment.RecordCount) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
        rsAdjustment.MoveNext
    Loop
    labProcessing.Caption = ""
    DoEvents
    Screen.MousePointer = 0
End Sub
