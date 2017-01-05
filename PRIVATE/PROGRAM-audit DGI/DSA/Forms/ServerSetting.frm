VERSION 5.00
Begin VB.Form frmServerSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Setting"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "ServerSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   390
      Top             =   5280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6450
      TabIndex        =   2
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4290
      TabIndex        =   1
      Top             =   5190
      Width           =   1035
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5370
      TabIndex        =   0
      Top             =   5190
      Width           =   1035
   End
   Begin VB.PictureBox picStep3 
      BorderStyle     =   0  'None
      Height          =   5025
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   7575
      TabIndex        =   19
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtRpt_PMIS 
         Height          =   375
         Left            =   3600
         TabIndex        =   48
         Top             =   4170
         Width           =   3195
      End
      Begin VB.CommandButton Command8 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   47
         Top             =   4170
         Width           =   345
      End
      Begin VB.TextBox txtRpt_OSMS 
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Top             =   3345
         Width           =   3195
      End
      Begin VB.CommandButton Command7 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   43
         Top             =   3360
         Width           =   345
      End
      Begin VB.CommandButton Command9 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   42
         Top             =   3750
         Width           =   345
      End
      Begin VB.CommandButton Command6 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   41
         Top             =   2970
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   40
         Top             =   2550
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   39
         Top             =   2145
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   38
         Top             =   1725
         Width           =   345
      End
      Begin VB.TextBox txtRpt_SMIS 
         Height          =   375
         Left            =   3600
         TabIndex        =   37
         Top             =   3750
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_HRMS 
         Height          =   375
         Left            =   3600
         TabIndex        =   36
         Top             =   2955
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_CSMS 
         Height          =   375
         Left            =   3600
         TabIndex        =   35
         Top             =   2535
         Width           =   3195
      End
      Begin VB.CommandButton Command1 
         Caption         =   "::"
         Height          =   405
         Left            =   6840
         TabIndex        =   31
         Top             =   1245
         Width           =   345
      End
      Begin VB.TextBox txtRpt_CRIS 
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   2130
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_AMIS 
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   1260
         Width           =   3195
      End
      Begin VB.TextBox txtRpt_CMIS 
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   1710
         Width           =   3195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PMIS"
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
         Left            =   3060
         TabIndex        =   49
         ToolTipText     =   "Sales Management Information System"
         Top             =   4260
         Width           =   405
      End
      Begin VB.Image Image4 
         Height          =   4125
         Left            =   0
         Picture         =   "ServerSetting.frx":08CA
         Top             =   900
         Width           =   2820
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OSMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   45
         ToolTipText     =   "Human Resource Management System"
         Top             =   3420
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   34
         ToolTipText     =   "Sales Management Information System"
         Top             =   3810
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HRMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   33
         ToolTipText     =   "Human Resource Management System"
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   32
         ToolTipText     =   "Customer Relation Information System"
         Top             =   2205
         Width           =   375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMIS 2.0 Report Path Configuration"
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
         TabIndex        =   26
         Top             =   240
         Width           =   3360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   25
         ToolTipText     =   "Accounting Management Information System"
         Top             =   1335
         Width           =   435
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   24
         ToolTipText     =   "Cash Monitoring Information System"
         Top             =   1785
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   23
         ToolTipText     =   "Car Service Management System"
         Top             =   2610
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   885
         Left            =   0
         Picture         =   "ServerSetting.frx":8070
         Top             =   0
         Width           =   7665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   165
      Left            =   90
      TabIndex        =   46
      Top             =   4980
      Width           =   7395
   End
   Begin VB.PictureBox picStep1 
      BorderStyle     =   0  'None
      Height          =   4995
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         Height          =   165
         Left            =   3570
         TabIndex        =   14
         Top             =   960
         Width           =   3435
      End
      Begin VB.PictureBox picConnection 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   6600
         ScaleHeight     =   945
         ScaleWidth      =   1095
         TabIndex        =   5
         Top             =   60
         Width           =   1095
         Begin VB.Image img4 
            Height          =   720
            Left            =   0
            Picture         =   "ServerSetting.frx":8B82
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image img3 
            Height          =   720
            Left            =   0
            Picture         =   "ServerSetting.frx":930F
            Top             =   0
            Width           =   720
         End
         Begin VB.Image img2 
            Height          =   720
            Left            =   0
            Picture         =   "ServerSetting.frx":9A7C
            Top             =   0
            Width           =   720
         End
         Begin VB.Image img1 
            Height          =   720
            Left            =   0
            Picture         =   "ServerSetting.frx":A1EB
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   4965
         Left            =   0
         Picture         =   "ServerSetting.frx":A99C
         ScaleHeight     =   4965
         ScaleWidth      =   3435
         TabIndex        =   4
         Top             =   -60
         Width           =   3435
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   3450
         ScaleHeight     =   3525
         ScaleWidth      =   3900
         TabIndex        =   6
         Top             =   1260
         Width           =   3900
         Begin VB.TextBox txtSQLServerName 
            Height          =   345
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   7
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtDATABASE 
            Height          =   345
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   9
            Top             =   2430
            Width           =   2175
         End
         Begin VB.TextBox txtSERVERNAME 
            Height          =   405
            Left            =   1590
            MaxLength       =   30
            TabIndex        =   11
            Top             =   2850
            Width           =   2175
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   525
            TabIndex        =   30
            Top             =   1290
            Width           =   135
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   525
            TabIndex        =   29
            Top             =   1050
            Width           =   135
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   525
            TabIndex        =   28
            Top             =   810
            Width           =   135
         End
         Begin VB.Label Label8 
            Caption         =   "Set up your Credentials"
            Height          =   225
            Left            =   720
            TabIndex        =   18
            Top             =   1290
            Width           =   3405
         End
         Begin VB.Label Label7 
            Caption         =   "Configure DMIS 2.0 Report Settings"
            Height          =   225
            Left            =   720
            TabIndex        =   17
            Top             =   1050
            Width           =   3405
         End
         Begin VB.Label Label6 
            Caption         =   "Configure DMIS 2.0 Server Settings"
            Height          =   225
            Left            =   720
            TabIndex        =   16
            Top             =   810
            Width           =   3405
         End
         Begin VB.Label Label5 
            Caption         =   "This Wizard Helps you to"
            Height          =   255
            Left            =   300
            TabIndex        =   15
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Server Name"
            Height          =   210
            Left            =   600
            TabIndex        =   13
            Top             =   2910
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Base"
            Height          =   210
            Left            =   720
            TabIndex        =   12
            Top             =   2460
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SQL Server Name"
            Height          =   210
            Left            =   180
            TabIndex        =   10
            Top             =   2100
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Welcome To DMIS 2.0 Server Setting"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   3510
         End
      End
      Begin VB.Label laberror 
         Height          =   765
         Left            =   3480
         TabIndex        =   27
         Top             =   180
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmServerSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BrowseInfo
    hwndOwner                                          As Long
    pIDLRoot                                           As Long
    pszDisplayName                                     As Long
    lpszTitle                                          As Long
    ulFlags                                            As Long
    lpfnCallback                                       As Long
    lParam                                             As Long
    iImage                                             As Long
End Type
Public ShowLogin                                       As Boolean
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'''DSN CREATOR
'Constant Declaration
Private Const ODBC_ADD_DSN = 1        ' Add data source
Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3     ' Remove data source
Private Const vbAPINull                                As Long = 0  ' NULL Pointer
'Function Declare
#If Win32 Then
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                                             (ByVal hwndParent As Long, ByVal fRequest As Long, _
                                              ByVal lpszDriver As String, ByVal lpszAttributes As String) _
                                              As Long
#Else
Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" _
                                             (ByVal hwndParent As Integer, ByVal fRequest As Integer, ByVal _
                                                                                                      lpszDriver As String, ByVal lpszAttributes As String) As Integer
#End If

'LOCAL VAIRABLE
Public intSteps                                        As Integer
Dim cn                                                 As ADODB.Connection
Dim userexists                                         As Boolean

'Upating Code       : AXP-0713200715:24
Private Sub cmdNext_Click()
''||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

    If intSteps = 0 Then
        cmdNext.Caption = "Next"
        laberror = ""

        Screen.MousePointer = 11
        If txtDATABASE = "" Then
            MessagePop InfoHelp, "Missing Field", "Database Name Missing"
            txtDATABASE.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        If txtSERVERNAME = "" Then
            MessagePop InfoHelp, "Missing Field", "Server Name Missing"
            txtSERVERNAME.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        If txtSQLServerName = "" Then
            MessagePop InfoHelp, "Missing Field", "SQL Server Name Missing"
            txtSQLServerName.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        On Error GoTo adder:
        If cn.State = 1 Then
            cn.Close
        End If
        cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & txtDATABASE & ";Data Source=" & txtSERVERNAME
        
        
        cn.Open
        DoEvents
        Timer1.Enabled = True
        Screen.MousePointer = 0
        intSteps = intSteps + 1
        picStep1.Visible = False
        picStep3.Visible = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdNext.Caption = "Finish"
        Savedbsettings

        txtRpt_AMIS = GetSetting("DMIS 2.0", "REPORTS", "AMIS", "\\" & ServerName & "\DMIS 2.0\REPORTS\AMIS")
        txtRpt_CMIS = GetSetting("DMIS 2.0", "REPORTS", "CMIS", "\\" & ServerName & "\DMIS 2.0\REPORTS\CMIS")
        txtRpt_CRIS = GetSetting("DMIS 2.0", "REPORTS", "CRIS", "\\" & ServerName & "\DMIS 2.0\REPORTS\CRIS")
        txtRpt_CSMS = GetSetting("DMIS 2.0", "REPORTS", "CSMS", "\\" & ServerName & "\DMIS 2.0\REPORTS\CSMS")
        txtRpt_HRMS = GetSetting("DMIS 2.0", "REPORTS", "HRMS", "\\" & ServerName & "\DMIS 2.0\REPORTS\HRMS")
        txtRpt_OSMS = GetSetting("DMIS 2.0", "REPORTS", "OSMS", "\\" & ServerName & "\DMIS 2.0\REPORTS\OSMS")
        txtRpt_SMIS = GetSetting("DMIS 2.0", "REPORTS", "SMIS", "\\" & ServerName & "\DMIS 2.0\REPORTS\SMIS")
        txtRpt_PMIS = GetSetting("DMIS 2.0", "REPORTS", "PMIS", "\\" & ServerName & "\DMIS 2.0\REPORTS\PMIS")


        If txtRpt_AMIS = "" Then
            txtRpt_AMIS = "\\" & ServerName & "\DMIS 2.0\REPORTS\AMIS"
        End If
        If txtRpt_CMIS = "" Then
            txtRpt_CMIS = "\\" & ServerName & "\DMIS 2.0\REPORTS\CMIS"
        End If
        If txtRpt_CRIS = "" Then
            txtRpt_CRIS = "\\" & ServerName & "\DMIS 2.0\REPORTS\CRIS"
        End If
        If txtRpt_CSMS = "" Then
            txtRpt_CSMS = "\\" & ServerName & "\DMIS 2.0\REPORTS\CSMS"
        End If

        If txtRpt_HRMS = "" Then
            txtRpt_HRMS = "\\" & ServerName & "\DMIS 2.0\REPORTS\HRMS"
        End If
        If txtRpt_OSMS = "" Then
            txtRpt_OSMS = "\\" & ServerName & "\DMIS 2.0\REPORTS\OSMS"
        End If
        If txtRpt_SMIS = "" Then
            txtRpt_SMIS = "\\" & ServerName & "\DMIS 2.0\REPORTS\SMIS"
        End If
        If txtRpt_PMIS = "" Then
        txtRpt_PMIS = "\\" & ServerName & "\DMIS 2.0\REPORTS\PMIS"
        End If
        

        ''||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    ElseIf intSteps = 1 Then
        If LTrim(RTrim(txtRpt_AMIS)) = "" Then
            If MsgBox("Accounting Module (AMIS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbYes Then
                txtRpt_AMIS.SetFocus
                Exit Sub
            End If
        End If
        If LTrim(RTrim(txtRpt_CMIS)) = "" Then
            Call MsgBox("Cash Monitoring Module (CMIS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE)
            txtRpt_CMIS.SetFocus
            Exit Sub
        End If
        If LTrim(RTrim(txtRpt_CRIS)) = "" Then
            If MsgBox("Customer Relation Module (CRIS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbCancel Then
                txtRpt_CRIS.SetFocus
                Exit Sub
            End If
        End If
        If LTrim(RTrim(txtRpt_CSMS)) = "" Then
            If MsgBox("Car Service Module (CSMS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbCancel Then
                txtRpt_CSMS.SetFocus
                Exit Sub
            End If
        End If

        If LTrim(RTrim(txtRpt_HRMS)) = "" Then
            If MsgBox("Human Resource Module (HRMS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbCancel Then
                txtRpt_HRMS.SetFocus
                Exit Sub
            End If
        End If
        If LTrim(RTrim(txtRpt_OSMS)) = "" Then
            If MsgBox("Office Supplies Module (OSMS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbCancel Then
                txtRpt_OSMS.SetFocus
                Exit Sub
            End If
        End If
        If LTrim(RTrim(txtRpt_SMIS)) = "" Then
            If MsgBox("Car Sales Module (SMIS) Report Path has not been Configured . Are You Sure ?", vbQuestion + vbOKCancel, App.TITLE) = vbCancel Then
                txtRpt_SMIS.SetFocus
                Exit Sub
            End If
        End If


        AMIS_REPORT_PATH = txtRpt_AMIS
        CMIS_REPORT_PATH = txtRpt_CMIS
        CSMS_REPORT_PATH = txtRpt_CSMS
        CRIS_REPORT_PATH = txtRpt_CRIS
        HRMS_REPORT_PATH = txtRpt_HRMS
        SMIS_REPORT_PATH = txtRpt_SMIS
        OSMS_REPORT_PATH = txtRpt_OSMS
        PMIS_REPORT_PATH = txtRpt_PMIS

        SaveSetting "DMIS 2.0", "REPORTS", "AMIS", txtRpt_AMIS
        SaveSetting "DMIS 2.0", "REPORTS", "CMIS", txtRpt_CMIS
        SaveSetting "DMIS 2.0", "REPORTS", "CSMS", txtRpt_CSMS
        SaveSetting "DMIS 2.0", "REPORTS", "CRIS", txtRpt_CRIS
        SaveSetting "DMIS 2.0", "REPORTS", "HRMS", txtRpt_HRMS
        SaveSetting "DMIS 2.0", "REPORTS", "SMIS", txtRpt_SMIS
        SaveSetting "DMIS 2.0", "REPORTS", "OSMS", txtRpt_OSMS
        SaveSetting "DMIS 2.0", "REPORTS", "PMIS", txtRpt_PMIS

        If ShowLogin = True Then
            frmSecurity.Show
        End If
        Unload Me
    End If
    Exit Sub
adder:
    laberror = Err.Description
    Err.Clear
    Screen.MousePointer = 0
    img4.Visible = True
    Timer1.Enabled = False

End Sub

Private Sub cmdPrev_Click()
    cmdNext.Caption = "Next"
    If intSteps = 1 Then
        picStep1.Visible = True

        picStep3.Visible = False
        cmdPrev.Enabled = False: cmdNext.Enabled = True
        intSteps = 0
    ElseIf intSteps = 2 Then
        picStep1.Visible = False

        picStep3.Visible = False
        cmdPrev.Enabled = True: cmdNext.Enabled = True
        intSteps = 1
    End If

End Sub

Private Sub Command1_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For AMIS Report Path.")
    If strResFolder <> "" Then
            txtRpt_AMIS = UCase(strResFolder)
        
            AMIS_REPORT_PATH = UCase(strResFolder)

        'If LTrim(RTrim(txtRpt_CMIS)) = "" Then
            txtRpt_CMIS = UCase(Replace(strResFolder, "AMIS", "CMIS"))
        'End If
        'If LTrim(RTrim(txtRpt_CRIS)) = "" Then
            txtRpt_CRIS = UCase(Replace(strResFolder, "AMIS", "CRIS"))
        'End If
        'If LTrim(RTrim(txtRpt_CSMS)) = "" Then
            txtRpt_CSMS = UCase(Replace(strResFolder, "AMIS", "CSMS"))
        'End If
        'If LTrim(RTrim(txtRpt_HRMS)) = "" Then
            txtRpt_HRMS = UCase(Replace(strResFolder, "AMIS", "HRMS"))
        'End If
        'If LTrim(RTrim(txtRpt_OSMS)) = "" Then
            txtRpt_OSMS = UCase(Replace(strResFolder, "AMIS", "OSMS"))
        'End If
        'If LTrim(RTrim(txtRpt_SMIS)) = "" Then
            txtRpt_SMIS = UCase(Replace(strResFolder, "AMIS", "SMIS"))
            
            txtRpt_PMIS = UCase(Replace(strResFolder, "AMIS", "PMIS"))
        'End If
    End If

End Sub

Private Sub Command2_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For CMIS Report Path.")
    If strResFolder <> "" Then
        txtRpt_CMIS = UCase(strResFolder)
        CMIS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For CRIS Report Path.")
    If strResFolder <> "" Then
        txtRpt_CRIS = UCase(strResFolder)
        CRIS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Command5_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For CSMS Report Path.")
    If strResFolder <> "" Then
        txtRpt_CSMS = UCase(strResFolder)
        CSMS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Command6_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For HRMS Report Path.")
    If strResFolder <> "" Then
        txtRpt_HRMS = UCase(strResFolder)
        HRMS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Command7_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For OSMS Report Path.")
    If strResFolder <> "" Then
        txtRpt_OSMS = UCase(strResFolder)
        OSMS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub


Private Sub Command8_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For PMIS Report Path.")
    If strResFolder <> "" Then
        txtRpt_PMIS = UCase(strResFolder)
        PMIS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Command9_Click()
    Dim strResFolder                                   As String
    strResFolder = BrowseForFolder(hwnd, "Please Select a Folder For SMIS Report Path.")
    If strResFolder <> "" Then
        txtRpt_SMIS = UCase(strResFolder)
        SMIS_REPORT_PATH = UCase(strResFolder)
    End If
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    
    txtSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    txtSQLServerName = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    txtDATABASE = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    
    picStep1.Visible = True
    picStep3.Visible = False
    
    On Error GoTo adder1:
    
   ' If txtSERVERNAME <> "" And txtDATABASE <> "" And txtSERVERNAME <> "" Then
   '     cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & txtDATABASE & ";Data Source=" & txtSERVERNAME
   '     cn.Open
   '     DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Database & ";Data Source=" & ServerName
   '     If cn.State = 1 Then
   '         Timer1.Enabled = True
   '     End If
   ' End If
  


10:     If intSteps = 0 Then
        picStep1.Visible = True

        picStep3.Visible = False
        cmdNext.Enabled = True: cmdPrev.Enabled = False

    ElseIf intSteps = 1 Then
        picStep1.Visible = False
        picStep3.Visible = True
        cmdNext.Enabled = False: cmdPrev.Enabled = True
    End If
    Exit Sub

adder1:
    If Err.Number <> 3704 Then
        MsgBox Err.Description
    End If
    Err.Clear
    intSteps = 0
    Timer1.Enabled = False
    GoTo 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Exit Sub

    Dim frm                                            As Form
    Dim Cnt                                            As Integer
    For Each frm In Forms
        If frm.Visible = False Then
            Cnt = Cnt + 1
        End If
    Next

    If Cnt = 1 Then
        For Each frm In Forms
            Unload frm
        Next
    Else
        Cancel = 1
    End If

End Sub

Private Sub txtDATABASE_Change()
    laberror = ""
End Sub

Private Sub txtSERVERNAME_Change()
    laberror = ""
End Sub

Private Sub txtSQLServerName_Change()
    laberror = ""
End Sub

Private Sub txtSQLServerName_LostFocus()
    txtSERVERNAME.Text = txtSQLServerName
End Sub

Private Sub Timer1_Timer()

    If img1.Visible = True Then
        img1.Visible = False
        img2.Visible = True
        img3.Visible = False
        img4.Visible = False
    ElseIf img2.Visible = True Then
        img1.Visible = False
        img2.Visible = False
        img3.Visible = True
        img4.Visible = False
    ElseIf img3.Visible = True Then
        img1.Visible = True
        img2.Visible = False
        img3.Visible = False
        img4.Visible = False
    ElseIf img4.Visible = True Then
        img1.Visible = False
        img2.Visible = False
        img3.Visible = False
    End If

End Sub

Private Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String

'declare variables to be used
    Dim iNull                                          As Integer
    Dim lpIDList                                       As Long
    Dim lResult                                        As Long
    Dim sPath                                          As String
    Dim udtBI                                          As BrowseInfo

    'initialise variables
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Call the browse for folder API
    lpIDList = SHBrowseForFolder(udtBI)

    'get the resulting string path
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
    End If

    'If cancel was pressed, sPath = ""
    BrowseForFolder = sPath

End Function

Function CreateDSN(xSERVERNAME, xDSN, xDatabase) As Boolean


#If Win32 Then
    Dim intRet                                         As Long
#Else
    Dim intRet                                         As Integer
#End If

    Dim strDriver                                      As String
    Dim strAttributes                                  As String
    strDriver = "SQL Server"
    strAttributes = "SERVER=" & xSERVERNAME & Chr$(0)
    strAttributes = strAttributes & "DESCRIPTION=DMIS 2.0 DATABASE" & Chr$(0)
    strAttributes = strAttributes & "DSN=" & xDSN & Chr$(0)
    strAttributes = strAttributes & "DATABASE=" & xDatabase & Chr$(0)
    strAttributes = strAttributes & "Trusted_connection=Yes" & Chr$(0)
    strAttributes = strAttributes & "UseProcForPrepare=0" & Chr$(0)
    strAttributes = strAttributes & "AutoTranslate=0" & Chr$(0)
    strAttributes = strAttributes & "AnsiNPW=No" & Chr$(0)
    'To show dialog, use Form1.Hwnd instead of vbAPINull.
    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, _
                                 strDriver, strAttributes)
    If intRet Then
        CreateDSN = True
    Else
        CreateDSN = False
    End If
End Function



Sub Savedbsettings()
    SaveSetting "DMIS 2.0", "SETTINGS", "SERVERNAME", txtSERVERNAME
    SaveSetting "DMIS 2.0", "SETTINGS", "SQLSERVERNAME", txtSQLServerName
    SaveSetting "DMIS 2.0", "SETTINGS", "DATABASE", txtDATABASE
    DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & txtDATABASE & ";Data Source=" & txtSQLServerName
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT; Data Source=" & txtSQLServerName
    'If CreateDSN(ServerName, "DMIS", Database) = False Then
    '    MsgBox " ODBC For DMIS Failed Please Check Your Data Source, Configure Manually", vbInformation
    'End If

    'If CreateDSN(ServerName, "DMIS_AUDIT", "DMIS_AUDIT") = False Then
    '    MsgBox " ODBC For DMIS_AUDIT Failed Please Check Your Data Source, Configure Manually", vbInformation

    'End If
End Sub
