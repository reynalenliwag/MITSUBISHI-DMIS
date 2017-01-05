VERSION 5.00
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "COF080~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cash Monitoring Information System"
   ClientHeight    =   7440
   ClientLeft      =   1125
   ClientTop       =   1080
   ClientWidth     =   15240
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "Main.frx":030A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3960
      Top             =   3180
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3420
      Top             =   3840
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   4560
      Top             =   1980
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1967D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19997
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A2E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A607
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A921
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1AC3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1AF55
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B26F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B589
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B8A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BBBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C1F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C50B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C825
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CB3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D793
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7185
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Level"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Login Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:25 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock Status"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   2520
      Top             =   3825
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   2
      Width           =   140
      Height          =   270
      Animation       =   1
      AnimateDelay    =   125
      ShowDelay       =   2500
      BackgroundBitmap=   60
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   2940
      Top             =   3840
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      DesignerControls=   "Main.frx":1DAAD
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ApplyPatches()

End Sub

'Function Feature   : Reminder Module
'Date               : 06/26/2007
'Last Update        : 06/26/2007
'Database Update    : Added Table For Reminder Called Cris Reminders
'Who Updated        : AXP
'Upating Code       : AXP-062620071225
Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]"
    '& App.Revision & "]"
    
    Call ApplyThemes
    Call ConfigurePopUps
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Exit CMIS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim frm                                                       As Form
        For Each frm In Forms
            If Not (frm Is Nothing) Then
                Unload frm
            End If
        Next
        UnloadForm Me
    Else
        Cancel = 1
        frmMainMenu.Show
    End If
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.Id
            '''''''''''''''''''''''''''''''''''Files''''''''''''''''''''''''''''''''''''''''''''''''
        Case FILES_TRANSACTION
            If Module_Access(LOGID, "FILES TRANSACTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "A"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES TRANSACTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "TRANSACTION"
            frmCMISSBookEntry.Caption = "TRANSACTION CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_BANK
            If Module_Access(LOGID, "FILES BANK", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "B"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES BANK"
            frmCMISSBookEntry.labCODE.Caption = "FILES CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "BANK NAME"
            frmCMISSBookEntry.Caption = "BANK CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_BRANCH
            If Module_Access(LOGID, "FILES BRANCH", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "C"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES BRANCH"
            frmCMISSBookEntry.labCODE.Caption = "FILES BRANCH"
            frmCMISSBookEntry.labDESCNAME.Caption = "BRANCH NAME"
            frmCMISSBookEntry.Caption = "BRANCH CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_OTHERTRANSACTION
            If Module_Access(LOGID, "FILES OTHER TRANSACTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "D"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES OTHER TRANSACTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "OTHER TRAN."
            frmCMISSBookEntry.Caption = "OTHER TRANSACTION CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_PETTYCASHLTOTYPE
            If Module_Access(LOGID, "FILES PETTY LTO TYPE", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "E"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES PETTY LTO TYPE"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "PETTY/LTO"
            frmCMISSBookEntry.Caption = "PETTY CASH/L.T.O. MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_CHECKCLASSIFICATION
            If Module_Access(LOGID, "FILES CHECK CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "F"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES CHECK CLASSIFICATION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "CHECK CLASS"
            frmCMISSBookEntry.Caption = "CHECK CODE CLASSIFICATION MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_PAYMENTCLASSIFICATION
            If Module_Access(LOGID, "FILES PAYMENT CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "F"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES PAYMENT CLASSIFICATION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "CHECK CLASS"
            frmCMISSBookEntry.Caption = "CHECK CODE CLASSIFICATION MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_DEPARTMENT
            If Module_Access(LOGID, "FILES DEPARTMENT", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "H"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES DEPARTMENT"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "DEPT. NAME"
            frmCMISSBookEntry.Caption = "DEPARMENT CODE MAINTENANCE"
            frmCMISSBookEntry.Show
        Case FILES_EMPLOYEE
            If Module_Access(LOGID, "FILES EMPLOYEE", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "I"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES EMPLOYEE"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "EMP. NAME"
            frmCMISSBookEntry.Caption = "EMPLOYEES CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_REPLENISHMENTENTRY
            If Module_Access(LOGID, "FILES REPLENISHMENT ENTRY", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "J"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES REPLENISHMENT ENTRY"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "DESCRIPTION"
            frmCMISSBookEntry.Caption = "REPLENISHMENT CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_INSURANCECOMPANYLISTING
            If Module_Access(LOGID, "FILES INSURANCE COMPANY LISTING", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "K"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES INSURANCE COMPANY LISTING"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "COMPANY"
            frmCMISSBookEntry.Caption = "INSURANCE COMPANY CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_INTEROFFICECOLLECTION
            If Module_Access(LOGID, "FILES INTER OFFICE COLLECTION", "DATA ENTRY") = False Then Exit Sub
            BOOKTYPE = "L"
            On Error Resume Next
            Unload frmCMISSBookEntry
            frmCMISSBookEntry.LocalAcess = "FILES INTER OFFICE COLLECTION"
            frmCMISSBookEntry.labCODE.Caption = "CODE"
            frmCMISSBookEntry.labDESCNAME.Caption = "OFFICE"
            frmCMISSBookEntry.Caption = "INTER OFFICE CODE MAINTENANCE"
            frmCMISSBookEntry.Show
            
        Case FILES_CHARTSOFACCOUNT
            If Module_Access(LOGID, "FILES CHARTS OF ACCOUNT", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_CUSTOMERDEPOSIT
            If Module_Access(LOGID, "FILES CUSTOMER DEPOSIT", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_HIST_ORHISTORYFILES
            If Module_Access(LOGID, "FILES O.R. HISTORY FILES", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_HIST_BANKDEPOSITSHISTORYFILE
            If Module_Access(LOGID, "FILES BANK DEPOSITS HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_HIST_CASHENCASHMENTHISTORYFILE
            If Module_Access(LOGID, "FILES CASH ENCASHMENT HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_HIST_PETTYCASHHISTORYFILE
            If Module_Access(LOGID, "FILES PETTY CASH HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_HIST_LTOHISTORYFILE
            If Module_Access(LOGID, "FILES L.T.O. HISTORY FILE", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_PAIDAPP_REPAIRORDER
            If Module_Access(LOGID, "FILES PAID REPAIR ORDER", "DATA ENTRY") = False Then Exit Sub
        
        Case FILES_PAIDAPP_PERCUSTOMER
            If Module_Access(LOGID, "FILES PAID APP PER CUSTOMER", "DATA ENTRY") = False Then Exit Sub
            
        Case FILES_PAIDAPP_CASHANDCHARGEINVOICE
            If Module_Access(LOGID, "FILES PAIDAPP CASHANDCHARGEINVOICE", "DATA ENTRY") = False Then Exit Sub
            
        Case FILES_PARTICULARS
            If Module_Access(LOGID, "FILES PARTICULARS", "DATA ENTRY") = False Then Exit Sub
            'frmCMISParticulars.Show
            
        Case FILES_BANKINFORMATION
            If Module_Access(LOGID, "FILES BANKINFORMATION", "DATA ENTRY") = False Then Exit Sub
            'frmCMISBankInfo.Show
            
        Case FILES_SIGNATORIES
            If Module_Access(LOGID, "FILES SIGNATORIES", "DATA ENTRY") = False Then Exit Sub
            'frmCMISSignatories.Show
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''TRANSACTIONS''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case TRANS_ORE_ORWITHVAT, TOOL_ORWITHVAT
            If Module_Access(LOGID, "TRANSACTION O.R. WITH VAT", "TRANSACTION") = False Then Exit Sub
            Unload frmCMISOREntry
            OR_VAT_NONVAT = "VAT"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Official Receipt Data Entry [With VAT]"
            
        Case TRANS_ORE_NONVATOR, TOOL_NONVATOR
            If Module_Access(LOGID, "TRANSACTION O.R. WITH NON VAT", "TRANSACTION") = False Then Exit Sub
            Unload frmCMISOREntry
            OR_VAT_NONVAT = "NONVAT"
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Official Receipt Data Entry [NON VAT]"
            
        Case TRANS_ORE_OLDOFFICIALRECEIPTS
            If Module_Access(LOGID, "TRANSACTION O.R. OLD OFFICIAL RECEIPTS", "TRANSACTION") = False Then Exit Sub
            Unload frmCMISOREntry
            frmCMISOREntry.Show
            frmCMISOREntry.Caption = "Old Official Receipts"
            
        Case TRANS_FORCE_ORWITHVAT
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OR WITHVAT", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "VAT"
            frmCMISCancelOREntry.Show
            
        Case TRANS_FORCE_NONVATOR
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OR NON WITHVAT", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "NONVAT"
            frmCMISCancelOREntry.Show
            
        Case TRANS_FORCE_CARDWITHOR
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE CARDWITHOR", "TRANSACTION") = False Then Exit Sub
            CANCEL_OR_VAT_NONVAT = "CARD_OR"
            frmCMISCancelOREntry.Show
            
        Case TRANS_FORCE_OLDOFFICIALRECEIPTS
            If Module_Access(LOGID, "TRANSACTION O.R. FORCE OLD OFFICIAL RECEIPTS", "TRANSACTION") = False Then Exit Sub
            
        Case TRANS_PETTYCASHENTRY, TOOL_PETTYCASHENTRY
            If Module_Access(LOGID, "TRANSACTION PETTY CASH ENTRY", "TRANSACTION") = False Then Exit Sub
            frmCMISPettyCash.Show
            
        Case TRANS_LTOFUNDENTRY, TOOL_LTOFUNDENTRY
            If Module_Access(LOGID, "TRANSACTION LTO FUND ENTRY", "TRANSACTION") = False Then Exit Sub
            frmCMISLTOFUND.Show
            
        Case TRANS_BANKDEPOSIT, TOOL_BANKDEPOSIT
            If Module_Access(LOGID, "TRANSACTION BANKDEPOSIT", "TRANSACTION") = False Then Exit Sub
            frmCMISBankDeposit.Show
            
        Case TRANS_CHECKENCASHMENT, TOOL_CHECKENCASHMENT
            If Module_Access(LOGID, "TRANSACTION CHECK ENCASHMENT", "TRANSACTION") = False Then Exit Sub
            frmCMISCheckEncashment.Show
            
        Case TRANS_OFFICIALRECEIPTCUTOFFENTRY, TOOL_OFFICIALRECEIPTCUTOFFENTRY
            If Module_Access(LOGID, "TRANSACTION OFFICIAL RECEIPT CUT-OFF ENTRY", "TRANSACTION") = False Then Exit Sub
            frmCMISProcessCUTOFF.Show vbModal
            
        Case TRANS_CASHIERCASHCOUNT, TOOL_CASHIERCASHCOUNT
            If Module_Access(LOGID, "TRANSACTION CASHIER CASH COUNT", "TRANSACTION") = False Then Exit Sub
            frmCMISCashCount.Show
            
        Case TRANS_PETTYCASH
            If Module_Access(LOGID, "TRANSACTION PETTYCASH", "TRANSACTION") = False Then Exit Sub
            frmCMISPettyCash.Show
            ' frmCMISExpenses.Show
            'BG HERE
            '

            'frmCMISPettyCash.Show
            
        Case TRANS_ENCASHMENT
            If Module_Access(LOGID, "TRANSACTION ENCASHMENT", "TRANSACTION") = False Then Exit Sub
            ' frmCMISEncashment.Show
            
        Case TRANS_DEPOSITS
            If Module_Access(LOGID, "TRANSACTION DEPOSITS", "TRANSACTION") = False Then Exit Sub
            'frmCMISDeposits.Show
            
        Case TRANS_CASHCOUNT
            If Module_Access(LOGID, "TRANSACTION CASHCOUNT", "TRANSACTION") = False Then Exit Sub
            frmCMISCashCount.Show
            
        Case TRANS_ORSYSTEM
            If Module_Access(LOGID, "TRANSACTION ORSYSTEM", "TRANSACTION") = False Then Exit Sub
            'frmCMISORSystem.Show vbModal
            
            '''''''''''''''''''''''''''''''''''Maintainence''''''''''''''''''''''''''''''''''''''''''''''''
        Case MATAIN_SYSTEMCONFIGURATION
            If Module_Access(LOGID, "MAINTAIN SYSTEM CONFIGURATION", "SYSTEM") = False Then Exit Sub
            frmCMISProfile.Show
            
        Case MAINTAIN_PASSWORDMAINTENANCE
            frmAccMaintenance.Show
            
        Case MAINTAIN_ADV_EDITCASHPOSITION
            If Module_Access(LOGID, "MAINTAIN ADVANCED EDITCASHPOSITION", "SYSTEM") = False Then Exit Sub
            frmEDITViewCashPosition.Show
            
            '''''''''''''''''''''''''''''''''''Report''''''''''''''''''''''''''''''''''''''''''''''''
        Case REPORT_CASHINFLOWPERDATEOFRANGE, TOOL_CASHINFLOWPERDATEOFRANGE
            If Module_Access(LOGID, "REPORT CASH INFLOW PER DATE OF RANGE", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash In Flow Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
            
        Case REPORT_CUSTOMEROVERPAYMENTREPORT, TOOL_CUSTOMEROVERPAYMENTREPORT
            If Module_Access(LOGID, "REPORT CUSTOMER OVER PAYMENT REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Customer Over-Payment Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
            
        Case REPORT_CASHTALLYREPORT, TOOL_CASHTALLYREPORT
            If Module_Access(LOGID, "REPORT CASH TALLY REPORT", "REPORTS") = False Then Exit Sub
            frmCMISCutDate.Show
            
        Case REPORT_CRCARD_CARDLISTINGREPORTONHAND, TOOL_CREDITCARDPAYMENTREPORTS
            If Module_Access(LOGID, "REPORT CARD LISTING REPORT ON-HAND", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Card Listing Report On Hand"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
            
        Case REPORT_CRCARD_CARDBANKDEPOSITREPORT
            If Module_Access(LOGID, "REPORT CARD BANK DEPOSIT REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Credit Card Bank Deposit Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
            
        Case REPORT_DAILYTRANSMITTALREPORT, TOOL_DAILYTRANSMITTALREPORT
            If Module_Access(LOGID, "REPORT DAILY TRANSMITTAL REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Daily Transmittal Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_VIEWCASHPOSITION, TOOL_VIEWCASHPOSITION
            'If Module_Access(LOGID, "VIEW CASH POSITION", "INQUIRY") = False Then Exit Sub
            'frmViewCashPosition.Show vbModal

            If Module_Access(LOGID, "VIEW CASH POSITION", "INQUIRY") = False Then Exit Sub

            If Module_Access(LOGID, "TRANSACTION O.R. WITH VAT", "TRANSACTION") = False Then
                frmViewPettyCashPosition.Show
            Else
                If COMPANY_CODE = "HGC" Then
                    frmViewCollectionPosition.Show
                Else
                    frmViewCashPosition.Show vbModal
                End If
            End If

        Case REPORT_CASHREC_SUMMARYJOURNALREPORT
            If Module_Access(LOGID, "REPORT CASHRECIEPT SUMMARYJOURNALREPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash Receipts Journal Report - Summary"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_DETAILJOURNALREPORT
            If Module_Access(LOGID, "REPORT CASHRECIEPT DETAIL JOURNAL REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash Receipts Journal Report - Detail"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_APSUMMARYREPORT
            If Module_Access(LOGID, "REPORT CASHRECIEPT AP SUMMARY REPORT", "REPORTS") = False Then Exit Sub
        
        Case REPORT_PETTYCASHREPLENISHMENT
            If Module_Access(LOGID, "REPORT PETTY CASH REPLENISHMENT", "REPORTS") = False Then Exit Sub
        
        Case REPORT_CASHREC_CASHONHAND
            If Module_Access(LOGID, "REPORT CASHRECIEPT CASH ON HAND", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Cash On Hand Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_OUTPUTTAX
            If Module_Access(LOGID, "REPORT CASHRECIEPT OUT PUT TAX", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "OutPut Tax Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_CORPORATETAX
            If Module_Access(LOGID, "REPORT CASHRECIEPT CORPORATE TAX", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Corporate Tax Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_SALESDISCOUNT
            If Module_Access(LOGID, "REPORT CASHRECIEPT SALES DISCOUNT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Sales Discount Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_CASHREC_JOURNALWITHREFERENCECODE
            If Module_Access(LOGID, "REPORT CASHRECIEPT JOURNAL WITH REFERENCE CODE", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Journal with Reference Code"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_REPLENISH_LTOSUMMARYREPORT, TOOL_REPLENISHMENTSUMMARYREPORT
            If Module_Access(LOGID, "REPORT REPLENISH LTO SUMMARY REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "L.T.O. Summary Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_REPLENISH_PETTYCASHSUMMARYREPORT
            If Module_Access(LOGID, "REPORT REPLENISH PETTY CASH SUMMARY REPORT", "REPORTS") = False Then Exit Sub
            CMIS_Report_Range = "Petty Cash Summary Report"
            Unload frmCMISReportRange
            frmCMISReportRange.Show
        
        Case REPORT_SMOKETESTPAYMENTSSUMMARYREPORT
            If Module_Access(LOGID, "REPORT SMOKE TEST PAYMENTS SUMMARY REPORT", "REPORTS") = False Then Exit Sub

            '''''''''''''''''''''''''''''''''''Tools''''''''''''''''''''''''''''''''''''''''''''''''
        Case TOOL_SETDATETIME
            If Module_Access(LOGID, "TOOL SETDATETIME", "SYSTEM") = False Then Exit Sub
            'frmTOOLSSetDateTime.Show
            '''''''''''''''''''''''''''''''''''Windows''''''''''''''''''''''''''''''''''''''''''''''''
        
        Case WINDOW_ABOUT, TOOL_ABOUTTHEAUTHOR
            frmAbout.Show
        
        Case WINDOW_EXIT, TOOL_EXITSYSTEM
            Unload Me
        
        Case TOOL_DASHBOARD
            frmMainMenu.Show

    End Select
End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .EnableOffice2007Frame True
        .LoadDesignerBars
        '    .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .StatusBar.Visible = True
        .Options.SyncFloatingToolbars = True
    End With
    With SkinFramework1
        .LoadSkin App.Path & "\Royale.cjstyles", ""
        '.LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        .ApplyWindow Me.hwnd
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    Dim ToolTipContext                                                As ToolTipContext
    Set ToolTipContext = CommandBars1.ToolTipContext
    With ToolTipContext
        .ShowTitleAndDescription True, xtpToolTipIconInfo
        .SetMargin 2, 2, 2, 2
        .MaxTipWidth = 180
        If .IsBalloonStyleSupported Then
            .Style = xtpToolTipBalloon
        Else
            .Style = xtpToolTipOffice2007
        End If
        .ShowShadow = True
    End With

End Sub

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub ConfigurePopUps()
    Dim Item                                                          As PopupControlItem
    PopCntrl.RemoveAllItems
    'PopCntrl.Icons.AddIcons ImageManager.Icons
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    'PopCntrl.VisualTheme = xtpPopupThemeOffice2003
    'PopCntrl.SetSize 270, 140

    Set Item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    Item.Button = True
    Item.IconIndex = 899
    Item.Id = 707
    Item.Height = 20
    Item.Width = 20
    Item.CenterIcon
    Set Item = PopCntrl.AddItem(10, 10, 218, 30, vbNullString)
    Item.TextColor = RGB(15, 48, 145)
    Item.Bold = True
    Item.Font.Size = 10
    Item.Hyperlink = False
    Set Item = PopCntrl.AddItem(10, 32, 60, 50, vbNullString)
    Item.Height = 50
    Item.Width = 50
    Item.IconIndex = 0
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(62, 32, 260, 50, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.Height = 50
    Item.Id = 655
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.TextColor = RGB(190, 1, 1)
    Item.Height = 50
    Item.Font.Size = 7
    Item.Hyperlink = False
End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.Id = 707 Then
        PopCntrl.Close
    End If

End Sub
''''''''''''''END REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''

Private Sub Timer1_Timer()

    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            frmSMIS_Files_Reminders.Show
        End If
    End If
End Sub

