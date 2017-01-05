VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISAC_DataEntryTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Entry of Tag Numbers"
   ClientHeight    =   6750
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   7650
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_DataEntryTag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   7650
   Begin VB.PictureBox picTags2 
      Height          =   2145
      Left            =   2190
      ScaleHeight     =   2085
      ScaleWidth      =   3105
      TabIndex        =   4
      Top             =   1860
      Width           =   3165
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type status of tag number (e.g. U for unposted)"
         Top             =   810
         Width           =   405
      End
      Begin VB.TextBox txtEndTag 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Type ending series of tag number (e.g. 10,20,30, etc.)"
         Top             =   420
         Width           =   1425
      End
      Begin VB.TextBox txtStartTag 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Enter starting series of tag number (1,2,3, etc.)"
         Top             =   30
         Width           =   1425
      End
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
         Left            =   2280
         MouseIcon       =   "AC_DataEntryTag.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "AC_DataEntryTag.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel Entry"
         Top             =   1200
         Width           =   735
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
         Left            =   1560
         MouseIcon       =   "AC_DataEntryTag.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "AC_DataEntryTag.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Entry"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   900
         TabIndex        =   7
         Top             =   870
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Series"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Series"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   90
         Width           =   1545
      End
   End
   Begin wizButton.cmd cmdTags2 
      Height          =   2265
      Left            =   2130
      TabIndex        =   8
      Top             =   1800
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   3995
      TX              =   "F2 - Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "AC_DataEntryTag.frx":11FC
   End
   Begin MSFlexGridLib.MSFlexGrid grdTags 
      Height          =   5655
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      TextStyleFixed  =   3
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "AC_DataEntryTag.frx":1218
   End
   Begin VB.PictureBox picMatAdjust 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3930
      ScaleHeight     =   855
      ScaleWidth      =   3615
      TabIndex        =   11
      Top             =   5820
      Width           =   3615
      Begin VB.CommandButton cmdF6 
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
         Left            =   2880
         MouseIcon       =   "AC_DataEntryTag.frx":1532
         MousePointer    =   99  'Custom
         Picture         =   "AC_DataEntryTag.frx":1684
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
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
         Left            =   2160
         MouseIcon       =   "AC_DataEntryTag.frx":19EA
         MousePointer    =   99  'Custom
         Picture         =   "AC_DataEntryTag.frx":1B3C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
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
         Left            =   1440
         Picture         =   "AC_DataEntryTag.frx":1EA2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPMISAC_DataEntryTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTAGS                                                            As ADODB.Recordset

Sub rsRefresh()
    Set rsTAGS = New ADODB.Recordset
    rsTAGS.Open "Select * from tags  order by val(tag) asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    With grdTags
        .Row = 0
        .FormatString = "TAG NO.      | Status | Remarks                                                 | " & _
                        "Part No                  | Duplicate               "
    End With
End Sub

Sub FillGrid()
    Dim kcnt                                                          As Integer
    kcnt = 0
    If Not rsTAGS.EOF And Not rsTAGS.BOF Then
        Screen.MousePointer = 11
        rsTAGS.MoveFirst
        Do While Not rsTAGS.EOF
            kcnt = kcnt + 1
            grdTags.AddItem Null2String(rsTAGS!Tag) & Chr(9) & _
                            Null2String(rsTAGS!Status) & Chr(9) & _
                            Null2String(rsTAGS!remarks) & Chr(9) & _
                            Null2String(rsTAGS!partno) & Chr(9) & _
                            Null2String(rsTAGS!Duplicate)
            rsTAGS.MoveNext
        Loop
        If kcnt <> 0 Then grdTags.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdAdd_Click()

    Dim rsDupTags                                                     As ADODB.Recordset
    Dim LastTag                                                       As Long
    Set rsDupTags = New ADODB.Recordset
    rsDupTags.Open "select tag from tags  order by Val(tag) asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    If Not rsDupTags.EOF And Not rsDupTags.BOF Then
        rsDupTags.MoveLast
        LastTag = N2Str2IntZero(rsDupTags!Tag)
    End If
    txtStartTag.Text = LastTag + 1
    txtEndTag.Text = LastTag + 1
    txtStatus.Text = "U"
    cmdTags2.ZOrder 0
    picTags2.ZOrder 0
    On Error Resume Next
    txtStartTag.SetFocus
End Sub

Private Sub cmdCancel_Click()
    cmdTags2.ZOrder 1
    picTags2.ZOrder 1
End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    Dim StartT, EndT                                                  As Long
    StartT = NumericVal(txtStartTag.Text)
    EndT = NumericVal(txtEndTag.Text)

    If EndT >= StartT Then
        Dim rsDupTags                                                 As ADODB.Recordset
        Dim LastTag, LastID                                           As Long
        Set rsDupTags = New ADODB.Recordset
        rsDupTags.Open "select id,tag from tags  order by Val(tag) asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        If Not rsDupTags.EOF And Not rsDupTags.BOF Then
            rsDupTags.MoveLast
            LastTag = N2Str2IntZero(rsDupTags!Tag)
            LastID = N2Str2IntZero(rsDupTags!ID)
            If StartT < LastTag Then
                Screen.MousePointer = 0
                MsgSpeechBox "Error: Starting Tag Number Already Exist!"
                Exit Sub
            End If
        End If
    Else
        Screen.MousePointer = 0
        MsgSpeechBox "Error: Starting Tag Number must not be greater than Ending Tag Number!"
        Exit Sub
    End If

    Dim vtxtTag, vtxtStatus                                           As String
    Dim KIM                                                           As Integer
    vtxtStatus = N2Str2Null(txtStatus.Text)
    For KIM = StartT To EndT
        LastID = LastID + 1
        vtxtTag = N2Str2Null(KIM)
        gconINVENTORY.Execute "Insert into tags " & _
                              "(id,tag,status) values (" & LastID & ", " & vtxtTag & ", " & vtxtStatus & ")"
    Next
    cmdTags2.ZOrder 1
    picTags2.ZOrder 1
    rsRefresh
    cleargrid grdTags
    InitGrid
    FillGrid
    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            cmdAdd_Click
        Case vbKeyEscape
            cmdTags2.ZOrder 1
            picTags2.ZOrder 1
        Case vbKeyF5
            Unload Me
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cleargrid grdTags
    rsRefresh
    cmdTags2.ZOrder 1
    picTags2.ZOrder 1
    InitGrid
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub txtEndTag_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtStartTag_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

