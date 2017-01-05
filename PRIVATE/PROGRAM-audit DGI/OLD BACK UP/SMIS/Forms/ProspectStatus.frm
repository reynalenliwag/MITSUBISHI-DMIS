VERSION 5.00
Begin VB.Form frmSMIS_Files_ProspectStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Prospect Status"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProspectStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picVI 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      ScaleHeight     =   465
      ScaleWidth      =   3645
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   3645
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   10
         Top             =   0
         Width           =   2445
      End
      Begin VB.Label Label4 
         Caption         =   " VI NO#"
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
         Left            =   0
         TabIndex        =   11
         Top             =   60
         Width           =   1005
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   8
      Top             =   630
      Width           =   3675
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3210
      MouseIcon       =   "ProspectStatus.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "ProspectStatus.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   3450
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   1635
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1740
      Width           =   3705
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "OK"
      Height          =   615
      Left            =   2640
      MouseIcon       =   "ProspectStatus.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "ProspectStatus.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Set Reminders"
      Top             =   3450
      Width           =   585
   End
   Begin VB.Label labStatus 
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1470
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "New Status"
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
      Left            =   90
      TabIndex        =   2
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Current Status"
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
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Additional Note"
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
      Left            =   90
      TabIndex        =   4
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label lblCap 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   720
      Width           =   3315
   End
End
Attribute VB_Name = "frmSMIS_Files_ProspectStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PROSPECTID                                                     As Long

Function SetStatus(XXX)
    XXX = UCase(XXX)
    If XXX = "OPEN" Then
        SetStatus = "O"
    ElseIf XXX = "CLOSED" Then
        SetStatus = "C"
    ElseIf XXX = "INACTIVE" Then
        SetStatus = "I"
    Else
        SetStatus = "L"
    End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSnooze_Click()
    Dim STATUS

    STATUS = SetStatus(Combo1.Text)
    If STATUS = "" Then
        ShowIsRequiredMsg " Status. Please Select From The List"
        Exit Sub
    End If
    If picVI.Visible = True Then
        Dim rsSO                                                      As ADODB.Recordset
        If RTrim(LTrim(Text2)) = "" Then
            MsgBox "Closing Prospect Needs Sales Invoice Number. Please Indicate Vehicle Invoice Number", vbInformation
            Exit Sub
        End If

        Set rsSO = gconDMIS.Execute("select * from smis_salesorder where vi_no='" & Repleys(Text2) & "'")
        If rsSO.EOF Or rsSO.BOF Then
            MsgBox "Invalid Invoice Number. Please Input Valid Vehicle Invoice Number", vbInformation
            Exit Sub
        End If
    End If
    gconDMIS.Execute "update cris_prospects set invoiceno=" & N2Str2Null(Text2) & " , NOTES=" & N2Str2Null(Text1) & "   ,STATUS=" & N2Str2Null(STATUS) & " where  Prospectid=" & PROSPECTID
    gconDMIS.Execute "update cris_prospects set LogClosingDate=getdate() where  Prospectid=" & PROSPECTID
    LogAudit "E", "PROSPECT STATUS UPDATED ", "PROSPECTID:" & PROSPECTID & "STATUS:" & Combo1
    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    Unload Me
End Sub

Private Sub Combo1_Change()
    If Combo1.Text = "CLOSED" Then
        picVI.Visible = True
    Else
        picVI.Visible = False
    End If
End Sub

Private Sub Combo1_Click()
    Combo1_Change
End Sub

Private Sub Combo1_LostFocus()
    Combo1.ListIndex = SelectCombo(Combo1, Combo1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Dim RS                                                            As ADODB.Recordset
    Dim StatusProspect                                                As String
    Set RS = gconDMIS.Execute("Select STATUS ,NOTES ,INVOICENO from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)

    If Not RS.BOF Or Not RS.EOF Then

        StatusProspect = Null2String(RS!STATUS)
        Text1 = Null2String(RS!Notes)
        Text2 = Null2String(RS!invoiceno)
        If StatusProspect = "O" Then
            labStatus = "OPEN"

        ElseIf StatusProspect = "C" Then
            labStatus = "CLOSED"
        ElseIf StatusProspect = "I" Then
            labStatus = "INACTIVE"
        ElseIf StatusProspect = "L" Then
            labStatus = "LOST SALES"
        Else
            labStatus = "OPEN"
        End If
    End If

    Combo1.AddItem "OPEN"
    Combo1.AddItem "CLOSED"
    Combo1.AddItem "INACTIVE"
    Combo1.AddItem "LOST SALES"

End Sub

Private Sub Text2_LostFocus()
    Text2 = Format(Text2, "000000")
End Sub

