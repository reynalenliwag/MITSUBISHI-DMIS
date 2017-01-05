VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmToolsCompact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compact/Repair Database"
   ClientHeight    =   480
   ClientLeft      =   3060
   ClientTop       =   3060
   ClientWidth     =   2670
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ToolCompactData.frx":0000
   ScaleHeight     =   480
   ScaleWidth      =   2670
   Begin wizButton.cmd cmdOK 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      TX              =   "&Okey"
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
      MPTR            =   99
      MICON           =   "ToolCompactData.frx":2D3C
   End
   Begin wizButton.cmd cmdCancel 
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      TX              =   "&Cancel"
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
      MPTR            =   99
      MICON           =   "ToolCompactData.frx":2E9E
   End
End
Attribute VB_Name = "frmToolsCompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
UnloadForm Me
End Sub

Private Sub cmdOK_Click()
gconHRMS.Close
CompactRepairAccessDB HRMS_DATABASE_PATH, wizVar.DecryptAccess("62696Fd±´i°¸—¤²¡¢¡")
gconHRMS.Open
Unload Me
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
End Sub

'=====================================================================================
'COMPACT AND REPAIR DATABASE WITH PASSWORD PROTECTION IF PROVIDED
'=====================================================================================
'    Call CompactRepairAccessDB(MyDatabasePathAndFile, MyPassword)
Public Sub CompactRepairAccessDB(ByVal sDBFILE As String, _
            Optional sPassword As String = "")
Dim sDBPATH As String, sDBNAME As String, sDB As String, sDBtmp As String
sDBNAME = sDBFILE 'extrapulate the file name
Do While InStr(1, sDBNAME, "\") <> 0
        sDBNAME = Right(sDBNAME, Len(sDBNAME) - InStr(1, sDBNAME, "\"))
Loop
'get the path name only
sDBPATH = Left(sDBFILE, Len(sDBFILE) - Len(sDBNAME))

sDB = sDBPATH & sDBNAME
sDBtmp = sDBPATH & "tmp" & sDBNAME

'Call the statement to execute compact and repair...
If sPassword <> "" Then
        Call DBEngine.CompactDatabase(sDB, sDBtmp, dbLangGeneral, , ";pwd=" & sPassword)
Else
        Call DBEngine.CompactDatabase(sDB, sDBtmp)
End If
'wait for the app to finish
        DoEvents
'remove the uncompressed original
        Kill sDB
'rename the compressed file to the original to restore for other functions
        Name sDBtmp As sDB
MsgBoxXP "HRMS Database Successfully Compacted!"
End Sub


