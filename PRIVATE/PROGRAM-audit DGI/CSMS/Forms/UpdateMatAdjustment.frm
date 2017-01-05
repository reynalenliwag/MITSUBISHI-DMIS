VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "xpbutton.ocx"
Begin VB.Form frmCSMSUpdateMatAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Materials Adjustment File"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "UpdateMatAdjustment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5730
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
      Left            =   4875
      MouseIcon       =   "UpdateMatAdjustment.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMatAdjustment.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit Window"
      Top             =   705
      Width           =   705
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Update"
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
      Left            =   4185
      MouseIcon       =   "UpdateMatAdjustment.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateMatAdjustment.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Update Materials Adjustment File"
      Top             =   705
      Width           =   705
   End
   Begin MSMask.MaskEdBox txtTrandate 
      Height          =   345
      Left            =   2460
      TabIndex        =   0
      Top             =   1110
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   30
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
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateMatAdjustment.frx":0BAF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateMatAdjustment.frx":0BCB
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
         TabIndex        =   5
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   6
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
            MICON           =   "UpdateMatAdjustment.frx":0BE7
         End
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
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Transaction Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1140
      Width           =   2385
   End
End
Attribute VB_Name = "frmCSMSUpdateMatAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMatAdjust                                                       As ADODB.Recordset

Private Sub cmdCheck_Click()
    If IsDate(txtTranDate.Text) = False Then
        MsgSpeechBox "Please Input Valid Transaction Date!"
        Exit Sub
    End If

    Set rsMatAdjust = New ADODB.Recordset
    rsMatAdjust.Open "select * from CSMS_MatAdjust where status = 'N' AND LASTUPDATE = '" & CDate(txtTranDate.Text) & "' order by MatCde asc", gconDMIS
    If rsMatAdjust.EOF And rsMatAdjust.BOF Then
        MsgSpeechBox "Error: Adjustment File is Empty or Adjustments had been Posted already!"
        Exit Sub
    Else
        txtTranDate.Enabled = False
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
    txtTranDate.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Sub UpdateAdjustment()
    Dim I                                                             As Integer
    Dim vTrandate, vMatCde, vMatDsc                                   As String
    Dim vID, vTranQty                                                 As Integer
    Dim vAdd, vMinus                                                  As Double
    Dim VStatus                                                       As String

    On Error GoTo Errorcode

    rsMatAdjust.MoveFirst
    Screen.MousePointer = 11
    Me.Caption = "Updating Adjustment to Transactions Master File"
    DoEvents
    I = 0
    Do While Not rsMatAdjust.EOF
        vID = rsMatAdjust!ID
        vTrandate = N2Date2Null(txtTranDate.Text)
        vMatCde = N2Str2Null(rsMatAdjust!MATCDE)
        vMatDsc = N2Str2Null(rsMatAdjust!MatDsc)
        vMinus = N2Str2Zero(rsMatAdjust!minus)
        vAdd = N2Str2Zero(rsMatAdjust!Add)
        VStatus = "'N'"
        If vAdd <> 0 Then
            gconDMIS.Execute "Insert into CSMS_TdayTran " & _
                             "(trandate,trantype,MatCde,MatDsc,status,tranqty,tranno,itemno,in_out)" & _
                           " values (" & vTrandate & ", 'ADJ'," & vMatCde & ", " & vMatDsc & ", 'N'," & vAdd & ",'111111','1111','I')"
        Else
            gconDMIS.Execute "Insert into CSMS_TdayTran " & _
                             "(trandate,trantype,MatCde,MatDsc,status,tranqty,tranno,itemno,in_out)" & _
                           " values (" & vTrandate & ", 'ADJ'," & vMatCde & ", " & vMatDsc & ", 'N'," & vMinus & ",'000000','0000','O')"
        End If
        gconDMIS.Execute "update CSMS_MatAdjust set status = 'P' where id = " & vID
        DoEvents
        I = I + 1
        progCPB.Value = (I / rsMatAdjust.RecordCount) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
        rsMatAdjust.MoveNext
    Loop
    labProcessing.Caption = ""
    DoEvents
    Screen.MousePointer = 0

    Exit Sub

Errorcode:

    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSUpdateMatAdjustment = Nothing
    UnloadForm Me
End Sub

Private Sub txtTrandate_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
