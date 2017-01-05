VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPMISAC_INVMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000010&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptInventory 
      Left            =   3780
      Top             =   5220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComDlg.CommonDialog cmdDialogINV 
      Left            =   3375
      Top             =   5925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraDE 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   180
      TabIndex        =   16
      Top             =   2820
      Width           =   1635
      Begin MSForms.Label mDEPhyTicket 
         Height          =   1665
         Left            =   90
         TabIndex        =   7
         Top             =   1830
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Add/Edit Physical Count Ticket"
         Size            =   "2566;2937"
         MousePointer    =   99
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mDETagNumbers 
         Height          =   1695
         Left            =   90
         TabIndex        =   6
         Top             =   30
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Data Entry of Tag Numbers"
         Size            =   "2566;2990"
         MousePointer    =   99
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraIQ 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   2010
      TabIndex        =   17
      Top             =   60
      Width           =   1635
      Begin MSForms.Label mIQTagMaster 
         Height          =   1125
         Left            =   105
         TabIndex        =   10
         Top             =   2505
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Tag Master List"
         Size            =   "2566;1984"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mIQDispLedger 
         Height          =   990
         Left            =   90
         TabIndex        =   8
         Top             =   75
         Width           =   1515
         BackColor       =   16777215
         Caption         =   "Display Ledger File"
         Size            =   "2672;1746"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mIQDispTag 
         Height          =   1365
         Left            =   90
         TabIndex        =   9
         Top             =   1095
         Width           =   1515
         BackColor       =   16777215
         Caption         =   "Display Tag Master File by Part Number"
         Size            =   "2672;2408"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraPR 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   3690
      TabIndex        =   18
      Top             =   60
      Width           =   1635
      Begin MSForms.Label mPRGenLedger 
         Height          =   915
         Left            =   90
         TabIndex        =   21
         Top             =   1860
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Generate Ledger File"
         Size            =   "2566;1614"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         MouseIcon       =   "AC_InvMenu.frx":0000
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mPRConsPhyCnt 
         Height          =   990
         Left            =   90
         TabIndex        =   12
         Top             =   840
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Consolidate Physical Count"
         Size            =   "2566;1746"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mPRCreateCutOff 
         Height          =   735
         Left            =   90
         TabIndex        =   11
         Top             =   60
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Create Cut Off Master File"
         Size            =   "2566;1296"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mPRPostEdited 
         Height          =   945
         Left            =   90
         TabIndex        =   13
         Top             =   2820
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Check Cut Off  Balances"
         Size            =   "2566;1667"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraRP 
      BorderStyle     =   0  'None
      Caption         =   "---------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   5310
      TabIndex        =   19
      Top             =   60
      Width           =   1935
      Begin MSForms.Label mRPLedgerRep 
         Height          =   855
         Left            =   45
         TabIndex        =   22
         Top             =   120
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Ledger Report"
         Size            =   "2566;1508"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mRPUnaccPartNo 
         Height          =   885
         Left            =   75
         TabIndex        =   20
         Top             =   1950
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Unacc. Part No."
         Size            =   "2566;1561"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mRPUnaccTagNo 
         Height          =   855
         Left            =   75
         TabIndex        =   15
         Top             =   2850
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Unacct. Tag No."
         Size            =   "2566;1508"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label mRPVarianceRep 
         Height          =   855
         Left            =   75
         TabIndex        =   14
         Top             =   1050
         Width           =   1455
         BackColor       =   16777215
         Caption         =   "Variance Report"
         Size            =   "2566;1508"
         MousePointer    =   99
         BorderColor     =   16384
         BorderStyle     =   1
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   60
      TabIndex        =   5
      Top             =   -30
      Width           =   1875
      Begin wizButton.cmd cmdDataEntry 
         Height          =   495
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "DataEntry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_InvMenu.frx":0162
      End
      Begin wizButton.cmd cmdProcessing 
         Height          =   495
         Left            =   90
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Processing"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_InvMenu.frx":047C
      End
      Begin wizButton.cmd cmdReports 
         Height          =   495
         Left            =   90
         TabIndex        =   3
         Top             =   1710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Reports"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_InvMenu.frx":0796
      End
      Begin wizButton.cmd cmdInquiry 
         Height          =   495
         Left            =   90
         TabIndex        =   1
         Top             =   690
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         TX              =   "Inquiry"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_InvMenu.frx":0AB0
      End
      Begin wizButton.cmd cmdExit 
         Height          =   555
         Left            =   90
         TabIndex        =   4
         Top             =   2220
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         TX              =   "EXIT INVENTORY"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_InvMenu.frx":0DCA
      End
   End
End
Attribute VB_Name = "frmPMISAC_INVMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
'Purpose : To create USER DSN through VB code
'By      : Manish Kumar Pandey
''''''''''''''''''''''''''''''''''''''''''''''''''
'Put Following declaration in the form you want to create the dsn from.....
''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Const REG_DWORD = 4&
Private Const REG_SZ = 1                                     'Constant for a string variable type.
Private Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
                                      "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                                                       phkResult As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
                                       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                                                         ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
                                                                                                                      cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
                                     (ByVal hKey As Long) As Long
Dim FILNAME                                                           As String
Dim INVENTORY_REPORT_CONNECTION                                       As String

''''''''''''''''''''''''''''''''''''''''''''''''''
' And Now You need is just add commnd  button and compy following code to its click event.
''''''''''''''''''''''''''''''''''''''''''''''''''
'Don forget to customize the varialbles according to your project.....
''''
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SETODBC()
    Dim DataSourceName                                                As String
    Dim Description                                                   As String
    Dim DriverPath                                                    As String
    Dim DriverId                                                      As Long
    Dim DriverName                                                    As String
    Dim User                                                          As String
    Dim PWD                                                           As String

    Dim lResult                                                       As Long
    Dim hKeyHandle                                                    As Long
    Dim hKeyHandSub                                                   As Long
    Dim DBQ                                                           As String

    'Specify the DSN parameters.

    DataSourceName = "INVENTORY"
    DBQ = FILNAME
    Description = "PHYSICAL COUNT INVENTORY DATABASE"
    DriverPath = "E:\windows\System32\odbcjt32.dll"
    PWD = ""
    DriverId = 19
    User = "admin"
    DriverName = "Microsoft Access Driver (*.mdb)"

    'Create the new DSN key.

    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName, hKeyHandle)

    'Set the values of the new DSN key.

    lResult = RegSetValueEx(hKeyHandle, "DBQ", 0&, REG_SZ, _
                            ByVal DBQ, Len(DBQ))
    lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
                            ByVal Description, Len(Description))
    lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
                            ByVal DriverPath, Len(DriverPath))
    lResult = RegSetValueEx(hKeyHandle, "DriverID", 0&, REG_DWORD, _
                            25, 4)
    lResult = RegSetValueEx(hKeyHandle, "FIL", 0&, REG_SZ, _
                            ByVal "MS Access", 9)
    lResult = RegSetValueEx(hKeyHandle, "PWD", 0&, REG_SZ, _
                            ByVal PWD, Len(PWD))
    lResult = RegSetValueEx(hKeyHandle, "SafeTransactions", 0&, REG_DWORD, _
                            0, 4)
    lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, _
                            ByVal User, Len(User))
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Open a new key as follows
    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName & "\Engines\Jet", hKeyHandSub)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lResult = RegSetValueEx(hKeyHandSub, "ImplicitCommitSync", 0&, REG_SZ, _
                            ByVal "", 0)
    lResult = RegSetValueEx(hKeyHandSub, "MaxBufferSize", 0&, REG_DWORD, _
                            2048, 4)
    lResult = RegSetValueEx(hKeyHandSub, "PageTimeout", 0&, REG_DWORD, _
                            5, 4)
    lResult = RegSetValueEx(hKeyHandSub, "Threads", 0&, REG_DWORD, _
                            3, 4)
    lResult = RegSetValueEx(hKeyHandSub, "UserCommitSync", 0&, REG_SZ, _
                            ByVal "Yes", 3)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Close the new Sub key.
    lResult = RegCloseKey(hKeyHandSub)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Close the new DSN key.

    lResult = RegCloseKey(hKeyHandle)

    'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
    'Specify the new value.
    'Close the key.

    lResult = RegCreateKey(HKEY_CURRENT_USER, _
                           "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
    lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
                            ByVal DriverName, Len(DriverName))
    lResult = RegCloseKey(hKeyHandle)

    INVENTORY_REPORT_CONNECTION = "DSN=INVENTORY;UID=admin;PWD=;DSQ=" & FILNAME
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDataEntry_Click()

    cmdDataEntry.Top = 180
    fraDE.Visible = True
    fraIQ.Visible = False
    fraPR.Visible = False
    fraRP.Visible = False
    fraDE.Left = cmdDataEntry.Left + 80
    fraDE.Top = cmdDataEntry.Top + 500
    cmdInquiry.Top = fraDE.Height + 750
    cmdProcessing.Top = cmdInquiry.Top + 500
    cmdReports.Top = cmdProcessing.Top + 500
    cmdExit.Top = cmdReports.Top + 500
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInquiry_Click()

    cmdDataEntry.Top = 180
    cmdInquiry.Top = cmdDataEntry.Top + 500
    fraIQ.Visible = True
    fraDE.Visible = False
    fraPR.Visible = False
    fraRP.Visible = False
    fraIQ.Left = cmdInquiry.Left + 80
    fraIQ.Top = cmdInquiry.Top + 500
    cmdProcessing.Top = fraIQ.Height + 750 + 500
    cmdReports.Top = cmdProcessing.Top + 500
    cmdExit.Top = cmdReports.Top + 500
End Sub

Private Sub cmdProcessing_Click()
    cmdDataEntry.Top = 180
    cmdInquiry.Top = cmdDataEntry.Top + 500
    cmdProcessing.Top = cmdInquiry.Top + 500
    fraPR.Visible = True
    fraDE.Visible = False
    fraIQ.Visible = False
    fraRP.Visible = False
    fraPR.Left = cmdProcessing.Left + 80
    fraPR.Top = cmdProcessing.Top + 500
    cmdReports.Top = fraPR.Height + 750 + 500 + 500
    cmdExit.Top = cmdReports.Top + 500
End Sub

Private Sub cmdReports_Click()
    cmdDataEntry.Top = 180
    cmdInquiry.Top = cmdDataEntry.Top + 500
    cmdProcessing.Top = cmdInquiry.Top + 500
    cmdReports.Top = cmdProcessing.Top + 500
    fraRP.Visible = True
    fraDE.Visible = False
    fraIQ.Visible = False
    fraPR.Visible = False
    fraRP.Left = cmdReports.Left + 80
    fraRP.Top = cmdReports.Top + 500
    cmdExit.Top = fraPR.Height + 750 + 500 + 500 + 500
End Sub

Private Sub Form_Load()
    UpLeftMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Me.Width = 2050: Me.Height = 6870
    Me.Left = Me.Left - 100: Me.Top = Me.Top - 50
    cmdDataEntry.Top = 180
    fraDE.Left = cmdDataEntry.Left + 80
    fraDE.Top = cmdDataEntry.Top + 500
    cmdInquiry.Top = fraDE.Height + 750
    cmdProcessing.Top = cmdInquiry.Top + 500
    cmdReports.Top = cmdProcessing.Top + 500
    cmdExit.Top = cmdReports.Top + 500
    DoEvents
    On Error Resume Next
    Dim MYPATH, PAYLNAME                                              As String
    MYPATH = App.Path
    cmdDialogINV.Filter = "Access Files (*.MDB)|*.MDB"
    cmdDialogINV.FilterIndex = 1
    cmdDialogINV.DefaultExt = "MDB"
    cmdDialogINV.DialogTitle = "Open Inventory Database"
    PAYLNAME = cmdDialogINV.FileName
    If MYPATH <> "\" Then
        cmdDialogINV.FileName = MYPATH & "\" & cmdDialogINV.FileName
    End If
    If PAYLNAME = "" Then
        cmdDialogINV.FileName = "*.MDB"
    End If
    cmdDialogINV.Action = 1
    If Err = 32755 Then Exit Sub
    FILNAME = cmdDialogINV.FileName
    Dim CS                                                            As String
    CS = wizVar.DecryptAccess("50726F@§É¥¥èï_Å∂†d}oNvmbÜÑlëmîßâN∂•èµÆí•®m±Ö¶®ñ±±p")
    On Error GoTo ERRORCODE
    Set gconINVENTORY = New ADODB.Connection
    gconINVENTORY.ConnectionString = CS & FILNAME
    gconINVENTORY.Open
    If Err = 32755 Then Exit Sub
    SETODBC
    Exit Sub

ERRORCODE:
    ShowADOErrors gconINVENTORY
    On Error Resume Next
    MsgSpeechBox "Warning: Inventory Database is Invalid or Corrupted!" & vbCrLf & _
                 "Inventory Menu will be Unloaded... Contact EDP Immediately"
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gconINVENTORY.Close
    UnloadForm Me
End Sub

Private Sub mDEPhyTicket_Click()
    frmPMISAddPhyCntTicket.Show
End Sub

Private Sub mDEPhyTicket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_DE
    mDEPhyTicket.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mDETagNumbers_Click()
    Screen.MousePointer = 11
    frmPMISDataEntryTag.Show
    Screen.MousePointer = 0
End Sub

Private Sub mDETagNumbers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_DE
    mDETagNumbers.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mIQDispLedger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_IQ
    mIQDispLedger.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mIQDispTag_Click()
    mIQDispTag.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    frmPMISDataEntryTagByPartNo.Show
    Screen.MousePointer = 0
End Sub

Private Sub mIQDispTag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_IQ
    mIQDispTag.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mIQTagMaster_Click()
    mIQTagMaster.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub mIQTagMaster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_IQ
    mIQTagMaster.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mPRConsPhyCnt_Click()
    Screen.MousePointer = 11
    frmPMISCreateConsPhyCNT.Show
    Screen.MousePointer = 0
End Sub

Private Sub mPRConsPhyCnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_PR
    mPRConsPhyCnt.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mPRCreateCutOff_Click()
    Screen.MousePointer = 11
    frmPMISCreateCutOffMaster.Show
    Screen.MousePointer = 0
End Sub

Private Sub mPRCreateCutOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_PR
    mPRCreateCutOff.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mPRGenLedger_Click()
    Screen.MousePointer = 11
    frmPMISGenLedgerFile.Show
    Screen.MousePointer = 0
End Sub

Private Sub mPRGenLedger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_PR
    mPRGenLedger.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mPRPostEdited_Click()
    mPRPostEdited.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    frmPMISCutOffCheckPrevBal.Show
    Screen.MousePointer = 0
End Sub

Private Sub mPRPostEdited_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_PR
    mPRPostEdited.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mRPLedgerRep_Click()
    mRPLedgerRep.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub mRPLedgerRep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_RP
    mRPLedgerRep.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mRPUnaccPARTNO_Click()
    mRPUnaccPartNo.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnacctSTOCKNO.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Private Sub mRPUnaccPARTNO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_RP
    mRPUnaccPartNo.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mRPUnaccTagNo_Click()
    mRPUnaccTagNo.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnacctTagNo.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Private Sub mRPUnaccTagNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_RP
    mRPUnaccTagNo.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mRPVarianceRep_Click()
    mRPVarianceRep.SpecialEffect = fmSpecialEffectSunken
    Screen.MousePointer = 11
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "Variance.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Private Sub mRPVarianceRep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_RP
    mRPVarianceRep.SpecialEffect = fmSpecialEffectRaised
End Sub

