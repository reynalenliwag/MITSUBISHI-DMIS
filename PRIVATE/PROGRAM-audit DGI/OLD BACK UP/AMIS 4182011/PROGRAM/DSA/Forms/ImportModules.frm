VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmImportModules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Modules"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IMPORTMODULES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   -270
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picImport 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   0
      Width           =   6105
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   705
         Left            =   5250
         MouseIcon       =   "IMPORTMODULES.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "IMPORTMODULES.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin wizProgBar.Prg Prg1 
         Height          =   525
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   926
         Picture         =   "IMPORTMODULES.frx":0E67
         ForeColor       =   0
         BarPicture      =   "IMPORTMODULES.frx":0E83
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Process"
         Height          =   705
         Left            =   4530
         MouseIcon       =   "IMPORTMODULES.frx":0E9F
         MousePointer    =   99  'Custom
         Picture         =   "IMPORTMODULES.frx":0FF1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label labCPB 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   90
         TabIndex        =   7
         Top             =   0
         Width           =   5805
      End
   End
   Begin VB.PictureBox picOpen 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6165
      TabIndex        =   2
      Top             =   0
      Width           =   6165
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   435
         Left            =   5520
         TabIndex        =   8
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   150
         Width           =   4515
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   435
         Left            =   4890
         TabIndex        =   3
         Top             =   150
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmImportModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    picOpen.Visible = True
    picImport.Visible = False
End Sub

'Upating Code       : AXP-0713200715:24
Private Sub cmdCheck_Click()

    Dim oConAccess                      As ADODB.Connection
    Dim TEMPRS                          As ADODB.Recordset
    Dim Cnt                             As Long
    Dim tCount                          As Long
    On Error GoTo ErrorCode:

    Set oConAccess = New ADODB.Connection
    oConAccess.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text1 & ";Persist Security Info=False")
    Set TEMPRS = New ADODB.Recordset
    TEMPRS.CursorLocation = adUseClient
    Call TEMPRS.Open("Select * from ALL_Rams_Modules", oConAccess, adOpenDynamic, adLockReadOnly)
    tCount = TEMPRS.RecordCount

    Prg1.Max = tCount
    Cnt = 1

    gconDMIS.BeginTrans
    gconDMIS.Execute ("DELETE FROM ALL_RAMS_MODULES")
    While Not (TEMPRS.EOF)
        SQL = "INSERT INTO ALL_Rams_Modules  (MAINMODULENAME, DESCRIPTIONS, MODULE_TYPE, MODULEID) Values "
        SQL = SQL & "('" & TEMPRS!MAINMODULENAME & "',"
        SQL = SQL & "'" & UCase(TEMPRS!DESCRIPTIONS) & "'" & ","
        SQL = SQL & "'" & UCase(TEMPRS!MODULE_TYPE) & "'" & ","
        SQL = SQL & TEMPRS!ModuleID & ")"
        gconDMIS.Execute (SQL)
        Prg1.Value = Cnt
        labCPB = ((Cnt / tCount) * 100) & " %"
        Cnt = Cnt + 1
        TEMPRS.MoveNext
    Wend
    MessagePop InfoOk, "Modules Imported", " Modules Imported Sucessfully"
    gconDMIS.CommitTrans

    Unload Me

    Exit Sub
    gconDMIS.RollbackTrans
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()

    comDialog.DialogTitle = "Open File"
    comDialog.FILTER = "Microsoft Access Database|*.mdb|"
    comDialog.FilterIndex = 1
    comDialog.Flags = cdlOFNFileMustExist
    comDialog.CancelError = True
    On Error Resume Next
    comDialog.ShowOpen

    If Err Then
        Err.Clear
        Command2.Enabled = False
        Exit Sub
    End If
    ' Displays a message box.
    Text1 = comDialog.FileName
    Command2.Enabled = True






End Sub

Private Sub Command2_Click()
    picImport.Visible = True
    picOpen.Visible = False
End Sub

Private Sub Form_Load()
    picImport.Visible = False
    picOpen.Visible = True
End Sub

