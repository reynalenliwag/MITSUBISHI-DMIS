VERSION 5.00
Begin VB.Form frmPwdMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Maintenance"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5175
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   30
      MouseIcon       =   "frmUser.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   5025
      TabIndex        =   19
      Top             =   1860
      Width           =   5085
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4290
         Picture         =   "frmUser.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   3570
         Picture         =   "frmUser.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   2850
         Picture         =   "frmUser.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   2130
         Picture         =   "frmUser.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   1410
         Picture         =   "frmUser.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   690
         Picture         =   "frmUser.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   -30
         Picture         =   "frmUser.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   3630
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   20
      Top             =   1860
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   690
         Picture         =   "frmUser.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   0
         Picture         =   "frmUser.frx":1B5A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   30
      TabIndex        =   13
      Top             =   -30
      Width           =   5085
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "txtUserName"
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtPassWord 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   22
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "txtPassWord"
         Top             =   1020
         Width           =   3645
      End
      Begin VB.TextBox txtConfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   22
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "txtConfirm"
         Top             =   1440
         Width           =   3645
      End
      Begin VB.TextBox txtUserCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "txtUserName"
         Top             =   150
         Width           =   3645
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label labPassCode 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4710
         TabIndex        =   16
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "User Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label labUserID 
      Caption         =   "Label3"
      Height          =   375
      Left            =   180
      TabIndex        =   21
      Top             =   1860
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmPwdMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As ADODB.Recordset
Dim AddorEdit As String


Private Sub cmdAdd_Click()
    txtUserName.Enabled = True
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    Frame1.Caption = "Add A User"
    AddorEdit = "ADD"
    txtUserCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure?", "Delete Selected Record", XP_YesNo, msg_Question) = True Then
       gconOSMS.Execute "Delete from [User] where id = " & labUserID
       rsRefresh
       On Error Resume Next
       rsUser.MoveFirst
       StoreMemvars
    End If
End Sub

Private Sub cmdEdit_Click()
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Caption = "Edit A Record"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim UserName2Find As String
    UserName2Find = InputBox("Please Input Username or User Code to Find", "FindUser", txtUserName.Text)
    If UserName2Find <> "" Then
        If RecordFound(UserName2Find) = True Then
           StoreMemvars
        Else
           MsgBox "Can't Find " & UserName2Find
        End If
    End If
End Sub

Function RecordFound(AAA As Variant) As Boolean
Dim rsRecordFound As ADODB.Recordset
Set rsRecordFound = New ADODB.Recordset
Set rsRecordFound = rsUser.Clone
rsRecordFound.Find "username = '" & AAA & "'"
If Not rsRecordFound.EOF Then
   rsUser.Bookmark = rsRecordFound.Bookmark
   RecordFound = True
   Else
    Set rsRecordFound = New ADODB.Recordset
    Set rsRecordFound = rsUser.Clone
        rsRecordFound.Find "usercode = '" & AAA & "'"
   If Not rsRecordFound.EOF Then
        rsUser.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
   Else
      RecordFound = False
      End If
   End If
End Function

Private Sub cmdNext_Click()
    rsUser.MoveNext
    If rsUser.EOF Then
       rsUser.MoveLast
       MsgBoxXP "Last of Record", "Last Record", XP_OKOnly, msg_Information
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsUser.MovePrevious
    If rsUser.BOF Then
       rsUser.MoveFirst
       MsgBoxXP "Beginning of Record", "Beginning of Record", XP_OKOnly, msg_Information
    End If
    StoreMemvars
End Sub

Private Sub cmdSave_Click()
    If txtUserCode.Text = "" Then
       MsgBoxXP "User Code must have a value", "Warning", XP_OKOnly, msg_Critical
       txtUserCode.SetFocus
       Exit Sub
    End If
    If txtUserName.Text = "" Then
       MsgBoxXP "User Name must have a value", "Warning", XP_OKOnly, msg_Critical
       txtUserName.SetFocus
       Exit Sub
    End If
    
    If Frame1.Caption = "New" Then
       Dim rsUserDup As ADODB.Recordset
       Set rsUserDup = New ADODB.Recordset
       Set rsUserDup = rsUser.Clone
       rsUserDup.Find "username = '" & txtUserName.Text & "'"
       If Not rsUserDup.EOF Then
          MsgBoxXP "Username Already Exist!", "Warning", XP_OKOnly, msg_Critical
          txtUserName.SetFocus
          Exit Sub
       End If
    End If
    
    If txtPassWord.Text = "" Then
       MsgBoxXP "Password must have a value", "Warning", XP_OKOnly, msg_Critical
       txtPassWord.SetFocus
       Exit Sub
    End If
    If txtConfirm.Text = "" Then
       MsgBoxXP "Confirm Password must have a value", "Warning", XP_OKOnly, msg_Critical
       txtConfirm.SetFocus
       Exit Sub
    End If
    
    If txtConfirm.Text <> txtPassWord.Text Then
       MsgBoxXP "Passwords do not match!", "Warning", XP_OKOnly, msg_Critical
       txtPassWord.SetFocus
       Exit Sub
    End If
    
    If Len(txtUserCode.Text) <> 3 Then
       MsgBoxXP "Length of User Code must be 3", "Warning", XP_OKOnly, msg_Critical
       Exit Sub
    End If
    
         If AddorEdit = "ADD" Then
            gconOSMS.Execute "Insert into [User] " & _
                               "(usercode,username, userpass) Values ('" & txtUserCode.Text & "','" & txtUserName.Text & "', '" & txtPassWord.Text & "')"
         Else
            gconOSMS.Execute "Update [User] set " & _
                               "usercode = '" & txtUserCode.Text & "', " & _
                               "username = '" & txtUserName.Text & "', " & _
                               "userpass = '" & txtPassWord.Text & "'" & _
                               "where id = " & labUserID.Caption
         End If
         rsRefresh
         On Error Resume Next
            rsUser.Find "username = '" & (txtUserName.Text) & "'"
            cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    CenterMe Me
    
       Set rsUser = New ADODB.Recordset
           rsUser.Open "select * from [User]", gconOSMS, adOpenKeyset
       StoreMemvars
    DrawXPCtl Me
End Sub

Sub StoreMemvars()
     If Not (rsUser.EOF Or rsUser.BOF) Then
        labUserID.Caption = rsUser!ID
        txtUserCode.Text = (Null2String(rsUser!usercode))
        txtUserName.Text = (Null2String(rsUser!UserName))
        txtPassWord.Text = (Null2String(rsUser!UserPass))
        txtConfirm.Text = (Null2String(rsUser!UserPass))
     Else
        AddorEdit = "ADD"
        Frame1.Enabled = True
        initMemvars
        Picture1.Visible = False
        Picture2.Visible = True
     End If
End Sub

Sub initMemvars()
    txtUserCode.Text = ""
    txtUserName.Text = ""
    txtPassWord.Text = ""
    txtConfirm.Text = ""
End Sub

Sub rsRefresh()
    Set rsUser = New ADODB.Recordset
        rsUser.Open "select * from [User] order by username", gconOSMS, adOpenKeyset
End Sub


