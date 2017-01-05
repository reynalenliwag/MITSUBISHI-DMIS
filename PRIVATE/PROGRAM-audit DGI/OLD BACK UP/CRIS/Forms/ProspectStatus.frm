VERSION 5.00
Begin VB.Form frmSMIS_Files_ProspectStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Prospect Status"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
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
   ScaleHeight     =   3660
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   120
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
      Top             =   2970
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   1635
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1260
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
      Top             =   2970
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
      Top             =   990
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
Public ProspectID                       As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSnooze_Click()
    Dim status

    status = SetStatus(Combo1.Text)
    If status = "" Then
        ShowIsRequiredMsg " Status. Please Select From The List"
        Exit Sub
    End If
    gconDMIS.Execute "update cris_prospects set NOTES=" & N2Str2Null(Text1) & " , STATUS=" & N2Str2Null(status) & " where  Prospectid=" & ProspectID

    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()

    Dim rs                              As ADODB.Recordset
    Dim StatusProspect                  As String
    Set rs = gconDMIS.Execute("Select STATUS ,NOTES from CRIS_PROSPECTS WHERE PROSPECTID=" & ProspectID)

    If Not rs.BOF Or Not rs.EOF Then
        StatusProspect = Null2String(rs!status)
        Text1 = Null2String(rs!Notes)

        If StatusProspect = "O" Then
            labStatus = "OPEN"
        ElseIf StatusProspect = "C" Then
            labStatus = "CLOSED"
        ElseIf StatusProspect = "I" Then
            labStatus = "INACTIVE"
        Else
            labStatus = "OPEN"
        End If
    End If

    Combo1.AddItem "OPEN"
    Combo1.AddItem "CLOSED"
    Combo1.AddItem "INACTIVE"

End Sub

Function SetStatus(XXX)
    XXX = UCase(XXX)
    If XXX = "OPEN" Then
        SetStatus = "O"
    ElseIf XXX = "CLOSED" Then
        SetStatus = "C"
    ElseIf XXX = "INACTIVE" Then
        SetStatus = "I"
    End If
End Function

