VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmToolsZipFiles 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HRMS ZIP and Backup Database File..."
   ClientHeight    =   1665
   ClientLeft      =   4395
   ClientTop       =   2265
   ClientWidth     =   9855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "ZIPFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "ZIPFiles.frx":000C
   ScaleHeight     =   1665
   ScaleWidth      =   9855
   Begin wizButton.cmd cmdUnload 
      Height          =   405
      Left            =   1440
      TabIndex        =   6
      Top             =   180
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "ZIPFiles.frx":2D48
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   1335
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17330
            Text            =   "MAIN"
            TextSave        =   "MAIN"
            Key             =   "MAIN"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstLog 
      Height          =   1230
      Left            =   2310
      TabIndex        =   4
      Top             =   60
      Width           =   5805
   End
   Begin VB.CheckBox chkAddComment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &Comment"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   660
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox chkEncrypt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "En&crypt"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   990
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdSingleFile 
      Caption         =   "Files"
      Height          =   375
      Left            =   5430
      TabIndex        =   1
      Top             =   3480
      Width           =   1275
   End
   Begin VB.CommandButton cmdRecurse 
      Caption         =   "&Recurse..."
      Height          =   375
      Left            =   4110
      TabIndex        =   0
      Top             =   3480
      Width           =   1275
   End
   Begin wizButton.cmd cmdBackup 
      Height          =   405
      Left            =   150
      TabIndex        =   7
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      TX              =   "&Back Up Now!"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "ZIPFiles.frx":2EAA
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   8160
      Picture         =   "ZIPFiles.frx":300C
      Top             =   -60
      Width           =   1800
   End
End
Attribute VB_Name = "frmToolsZipFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_cZ As wizZIPClass
Attribute m_cZ.VB_VarHelpID = -1

'Private Sub cmdRecurse_Click()
'   With m_cZ
'      .ZipFile = App.Path & "\Test_Rec.zip"
'      .Encrypt = (chkEncrypt.Value = vbChecked)
'      .AddComment = (chkAddComment.Value = vbChecked)
'      .BasePath = App.Path
'      .ClearFileSpecs
'      .AddFileSpec "*.fr*"
'      .AddFileSpec "*.cls"
'      .AddFileSpec "*.bas"
'      .StoreFolderNames = True
'      .RecurseSubDirs = True
'      .Zip
'
'      If (.Success) Then
'         MsgBox "Zipped files." _
'            & vbCrLf & vbCrLf & _
'            "   Source: files matching *.fr*;*.cls;*.bas from " & .BasePath & vbCrLf & _
'            "   To: " & .ZipFile, vbInformation
'      Else
'         MsgBox "Zipping failed.", vbExclamation
'      End If
'   End With
'End Sub

Private Sub cmdBackup_Click()
cmdSingleFile_Click
End Sub

Private Sub cmdSingleFile_Click()
Dim cc As New wizCmnDialog
Dim sFIle As String

   ' Get the file to zip:
   'If (cc.VBGetOpenFileName(sFile, , , , , , "All Files (*.*)|*.*", , , "Choose File to Zip", , Me.hwnd)) Then
   sFIle = HRMS_DATABASE_PATH
      With m_cZ
         .Encrypt = (chkEncrypt.Value = vbChecked)
         .AddComment = (chkAddComment.Value = vbChecked)
         .ZipFile = HRMS_BACKUP_PATH & "HRMS_" & Month(LOGDATE) & Day(LOGDATE) & Year(LOGDATE) & ".zip"
         .StoreFolderNames = False
         .RecurseSubDirs = False
         .ClearFileSpecs
         .AddFileSpec sFIle
         .Zip
         If (.Success) Then
            MsgBox "Successfully Zipped and BACKUP Database File." _
               & vbCrLf & vbCrLf & _
               "   Source: " & .FileSpec(1) & vbCrLf & _
               "   To: " & .ZipFile, vbInformation
            cmdBackup.Enabled = False
         Else
            MsgBox "Zip Failed.", vbExclamation
         End If
      
      End With
   'End If
End Sub

Private Sub cmdUnload_Click()
Unload Me
End Sub

Private Sub Form_Load()
   Set m_cZ = New wizZIPClass
   CenterMe frmMain, Me, 1
End Sub

Private Sub m_cZ_CommentRequest(sComment As String, bCancel As Boolean)
   '
Dim sComm As String
   'sComm = InputBox("Enter comment:", App.EXEName)
   sComm = "HRMS DATABASE BACKUP AS OF: " & Date
   sComm = Trim(sComm)
   If Len(sComm) = 0 Then
      bCancel = True
   Else
      sComment = sComm
   End If
   '
End Sub

Private Sub m_cZ_PasswordRequest(sPassword As String, ByVal lMaxPasswordLength As Long, ByVal bConfirm As Boolean, bCancel As Boolean)
   '
Dim sPass As String
Dim sMsg As String
   If (bConfirm) Then
      sMsg = "Confirm password:"
   Else
      sMsg = "Enter password:"
   End If
   'sPass = InputBox(sMsg, App.EXEName)
   sPass = wizVar.DecryptAccess("62696Fd±´i°¸—¤²¡¢¡")
   sPass = Trim(sPass)
   If (Len(sPass) = 0) Then
      bCancel = True
   Else
      sPassword = sPass
   End If
   
   '
End Sub

Private Sub m_cZ_Progress(ByVal lCount As Long, ByVal sMsg As String)
   sbrMain.Panels(1).Text = sMsg
   lstLog.AddItem sMsg
   lstLog.ListIndex = lstLog.NewIndex
End Sub
