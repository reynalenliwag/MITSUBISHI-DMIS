VERSION 5.00
Begin VB.Form frmFiles_CopySettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Settings"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FILES_COPYSETTINGS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picStep2 
      BorderStyle     =   0  'None
      Height          =   5025
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox cboCopyTo 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2490
         Width           =   2835
      End
      Begin VB.ComboBox cboCopyFrom 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   2835
      End
      Begin VB.Image Image3 
         Height          =   4125
         Left            =   0
         Picture         =   "FILES_COPYSETTINGS.frx":08CA
         Top             =   900
         Width           =   2820
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Users Right Access Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   885
         Left            =   0
         Picture         =   "FILES_COPYSETTINGS.frx":4FA8
         Top             =   0
         Width           =   7665
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Your User Name and Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Width           =   4185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Copy Setting From User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   2
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Copy Setting To User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   1
         Top             =   2220
         Width           =   2250
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   6420
      TabIndex        =   5
      Top             =   5160
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   405
      Left            =   5400
      TabIndex        =   4
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   160
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   7575
   End
End
Attribute VB_Name = "frmFiles_CopySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCopyFrom_Click()
    Dim i                               As Integer
    cboCopyTo.Clear
    For i = 0 To cboCopyFrom.ListCount - 1
        If cboCopyFrom.List(i) <> cboCopyFrom.Text Then
            cboCopyTo.AddItem cboCopyFrom.List(i)
            cboCopyTo.ItemData(cboCopyTo.NewIndex) = cboCopyFrom.ItemData(i)
        End If
    Next
End Sub

'Upating Code       : AXP-0713200715:21
'---------------------------------------------------------------------------------------
' Procedure : Command1_Click
' DateTime  : 10/31/2007 10:17
' Author    : Ashish
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Command1_Click()

    Dim COPYTO                          As Integer
    Dim COPYFROM                        As Integer
    'On Error GoTo Errorcode:

    On Error GoTo ErrorCode                                  'AXP063110:17

    Screen.MousePointer = 11
    If cboCopyFrom.ListIndex = -1 Then
        MsgBox "Entry Must Match an Item in the List", vbCritical
        cboCopyFrom.SetFocus
        Exit Sub
    End If
    If cboCopyTo.ListIndex = -1 Then
        MsgBox "Entry Must Match an Item in the List", vbCritical
        cboCopyTo.SetFocus
        Exit Sub
    End If

    If MsgBox("Your Copying Setting From " & cboCopyFrom & " To " & cboCopyTo & vbCrLf & _
              "All the Acess Rights of " & cboCopyFrom & " will be Copied to " & cboCopyTo & vbCrLf & " Are You Sure?", vbQuestion + vbYesNo, App.TITLE) = vbNo Then

        Exit Sub

    End If


    COPYTO = cboCopyTo.ItemData(cboCopyTo.ListIndex)
    COPYFROM = cboCopyFrom.ItemData(cboCopyFrom.ListIndex)
    Dim SQL                             As String

    gconDMIS.Execute "Delete from ALL_Rams_User_Modules where USERID=" & COPYTO
    gconDMIS.Execute ("INSERT INTO ALL_Rams_User_Modules  Select " & COPYTO & ",  MAINMODULEID,MAINMODULENAME From ALL_Rams_User_Modules WHERE USERID=" & COPYFROM)

    gconDMIS.Execute "DELETE FROM ALL_Rams_UsersAcess     WHERE     (USERID =" & COPYTO & ")"

    SQL = "Insert Into ALL_Rams_UsersAcess (MODULEID,USERID,Acess_Add, Acess_Edit, Acess_Delete, Acess_View, Acess_Print, Acess_Process, Acess_System, Acess_Post, Acess_UnPost, Acess_CancelEntry ,Acess_Reprint,Acess_Export,Acess_Detail)"
    SQL = SQL & " SELECT  MODULEID, " & COPYTO & ", Acess_Add, Acess_Edit, Acess_Delete, Acess_View, Acess_Print, Acess_Process, Acess_System, Acess_Post, Acess_UnPost, Acess_CancelEntry ,Acess_Reprint,Acess_Export,Acess_Detail From ALL_Rams_UsersAcess Where (USERID = " & COPYFROM & ")"

    gconDMIS.Execute SQL
    Screen.MousePointer = 0



    Unload Me
    'Exit Sub
    'Errorcode:
    'ShowVBError


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    Dim TEMPRS                          As ADODB.Recordset
    If CHANGE_USER = True Then
        Set TEMPRS = gconDMIS.Execute("SELECT  USER_NAME AS USERNAME,USERID FROM ALL_RAMS_USERS WHERE USERGROUP<>'SDM' order by user_name")
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT  USERNAME,USERID FROM ALL_RAMS_USERS WHERE USERGROUP<>'SDM' order by username")
    End If

    While Not TEMPRS.EOF
        cboCopyFrom.AddItem (Null2String(TEMPRS!Username))
        cboCopyFrom.ItemData(cboCopyFrom.NewIndex) = TEMPRS!UserID
        TEMPRS.MoveNext
    Wend


End Sub
