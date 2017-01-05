VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COC288~1.OCX"
Begin VB.Form frmPMISINVMenuNew 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   720
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0275
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":050E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0691
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0829
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":09D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0B73
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0E89
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":0F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":121C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":1328
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":166A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":180A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPMISINVMenuNew.frx":19A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport rptInventory 
      Left            =   405
      Top             =   0
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
      Left            =   0
      Top             =   705
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Align           =   3  'Align Left
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _Version        =   655364
      _ExtentX        =   4815
      _ExtentY        =   12832
      _StockProps     =   64
      VisualTheme     =   11
      Animation       =   2
      Behaviour       =   1
      SelectItemOnFocus=   -1  'True
      ItemLayout      =   3
      HotTrackStyle   =   2
      SingleSelection =   -1  'True
      ColumnWidth     =   200
      MinimumGroupClientHeight=   50
   End
End
Attribute VB_Name = "frmPMISINVMenuNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose : To create USER DSN through VB code
'By      : Manish Kumar Pandey
''''''''''''''''''''''''''''''''''''''''''''''''''
'Put Following declaration in the form you want to create the dsn from.....
Option Explicit
Private Const REG_DWORD = 4&
Private Const REG_SZ = 1                              'Constant for a string variable type.
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
Dim FILNAME                                            As String
Dim INVENTORY_REPORT_CONNECTION                        As String




Sub CreateTaskPanel()
    Dim Group                                          As TaskPanelGroup

    Set Group = wndTaskPanel.Groups.Add(0, "Data Entry")
    Group.Items.Add 100, "Data Entry of Tag Number", xtpTaskItemTypeLink, 1
    Group.Items.Add 101, "Add or Edit Physical Count Ticket", xtpTaskItemTypeLink, 1
    Group.Items.Add 1007, "Exit", xtpTaskItemTypeLink, 16

    Set Group = wndTaskPanel.Groups.Add(0, "Inquiry")
    Group.Items.Add 102, "Display Ledger File", xtpTaskItemTypeLink, 1
    Group.Items.Add 103, "Display Tag Master File By Part Number", xtpTaskItemTypeLink, 1
    Group.Items.Add 104, "Tag Master List", xtpTaskItemTypeLink, 1
    Group.Items.Add 1007, "Exit", xtpTaskItemTypeLink, 16

    Set Group = wndTaskPanel.Groups.Add(0, "Processing")
    Group.Items.Add 105, "Create Cut Off Master File", xtpTaskItemTypeLink, 1
    Group.Items.Add 106, "Consolidate Physical Count", xtpTaskItemTypeLink, 1
    Group.Items.Add 107, "General Ledger File", xtpTaskItemTypeLink, 1
    Group.Items.Add 108, "Check Cut Off Balance", xtpTaskItemTypeLink, 1
    Group.Items.Add 1007, "Exit", xtpTaskItemTypeLink, 16

    Set Group = wndTaskPanel.Groups.Add(0, "Reports")
    Group.Items.Add 109, "Ledger Report", xtpTaskItemTypeLink, 1
    Group.Items.Add 110, "Variance Report", xtpTaskItemTypeLink, 1
    Group.Items.Add 111, "Unaccounted Part No", xtpTaskItemTypeLink, 1
    Group.Items.Add 112, "Unaccounted Tag No", xtpTaskItemTypeLink, 1
    Group.Items.Add 1007, "Exit", xtpTaskItemTypeLink, 16

    wndTaskPanel.SetImageList imlTaskPanelIcons
    Call wndTaskPanel.SetGroupIconSize(16, 16)
    wndTaskPanel.Groups(1).Items(1).Selected = True




End Sub


Private Sub wndTaskPanel_ItemClick(ByVal Item As ITaskPanelGroupItem)

    Select Case Item.ID
        Case 100
            Screen.MousePointer = 11
            frmPMISDataEntryTag.Show
            Screen.MousePointer = 0
        Case 101
            frmPMISAddPhyCntTicket.Show
        Case 102
        Case 103
            Screen.MousePointer = 11
            frmPMISDataEntryTagByPartNo.Show
            Screen.MousePointer = 0
        Case 104

        Case 105
            Screen.MousePointer = 11
            frmPMISCreateCutOffMaster.Show
            Screen.MousePointer = 0
        Case 106
            Screen.MousePointer = 11
            frmPMISCreateConsPhyCNT.Show
            Screen.MousePointer = 0
        Case 107
            Screen.MousePointer = 11
            frmPMISGenLedgerFile.Show
            Screen.MousePointer = 0
        Case 108
            Screen.MousePointer = 11
            frmPMISCutOffCheckPrevBal.Show
            Screen.MousePointer = 0


        Case 109
        Case 110
            Screen.MousePointer = 11
            rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptInventory, PMIS_REPORT_PATH & "Variance.rpt", "", INVENTORY_REPORT_CONNECTION, 1
            Screen.MousePointer = 0

        Case 111
            Screen.MousePointer = 11
            rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnacctSTOCKNO.rpt", "", INVENTORY_REPORT_CONNECTION, 1
            Screen.MousePointer = 0
        Case 112
            Screen.MousePointer = 11
            rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnacctTagNo.rpt", "", INVENTORY_REPORT_CONNECTION, 1
            Screen.MousePointer = 0
        Case 1007
            Unload Me
    End Select
End Sub

' And Now You need is just add commnd  button and compy following code to its click event.
'Don forget to customize the varialbles according to your project.....

Sub SETODBC()
    Dim DataSourceName                                 As String
    Dim Description                                    As String
    Dim DriverPath                                     As String
    Dim DriverId                                       As Long
    Dim DriverName                                     As String
    Dim User                                           As String
    Dim PWD                                            As String

    Dim lResult                                        As Long
    Dim hKeyHandle                                     As Long
    Dim hKeyHandSub                                    As Long
    Dim DBQ                                            As String


    DataSourceName = "INVENTORY"
    DBQ = FILNAME
    Description = "PHYSICAL COUNT INVENTORY DATABASE"
    DriverPath = "E:\windows\System32\odbcjt32.dll"
    PWD = ""
    DriverId = 19
    User = "admin"
    DriverName = "Microsoft Access Driver (*.mdb)"


    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName, hKeyHandle)


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
    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName & "\Engines\Jet", hKeyHandSub)
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
    lResult = RegCloseKey(hKeyHandSub)


    lResult = RegCloseKey(hKeyHandle)


    lResult = RegCreateKey(HKEY_CURRENT_USER, _
                           "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
    lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
                            ByVal DriverName, Len(DriverName))
    lResult = RegCloseKey(hKeyHandle)

    INVENTORY_REPORT_CONNECTION = "DSN=INVENTORY;UID=admin;PWD=;DSQ=" & FILNAME
End Sub

Private Sub Form_Load()
    UpLeftMe frmMain, Me, 1
    Me.Left = Me.Left - 100: Me.Top = Me.Top - 50
    Dim MYPATH, PAYLNAME                               As String
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
    Dim CS                                             As String
    CS = wizVar.DecryptAccess("50726F@§É¥¥èï_Å∂†d}oNvmbÜÑlëmîßâN∂•èµÆí•®m±Ö¶®ñ±±p")
    On Error GoTo ErrorCode
    Set gconINVENTORY = New ADODB.Connection
    gconINVENTORY.ConnectionString = CS & FILNAME
    gconINVENTORY.Open
    CreateTaskPanel
    If Err = 32755 Then Exit Sub
    SETODBC
    Exit Sub

ErrorCode:
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



