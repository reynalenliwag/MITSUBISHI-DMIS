VERSION 5.00
Begin VB.Form frmSMIS_Files_Signatories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signatories"
   ClientHeight    =   5220
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Signatories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9165
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5715
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   2190
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   240
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Application/Forms"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   30
      TabIndex        =   3
      Top             =   660
      Width           =   9105
      Begin VB.TextBox txtGenManDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2940
         Width           =   3405
      End
      Begin VB.TextBox txtSDispatchDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1740
         Width           =   3405
      End
      Begin VB.TextBox txtDelByDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2160
         Width           =   3405
      End
      Begin VB.TextBox txtFinManDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   28
         Top             =   2550
         Width           =   3405
      End
      Begin VB.TextBox txtSManagerDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1320
         Width           =   3405
      End
      Begin VB.TextBox txtChkDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   26
         Top             =   900
         Width           =   3405
      End
      Begin VB.TextBox txtPreRevDesig 
         Height          =   375
         Left            =   5610
         MaxLength       =   50
         TabIndex        =   25
         Top             =   480
         Width           =   3405
      End
      Begin VB.TextBox txtSig_GM 
         Height          =   375
         Left            =   2130
         TabIndex        =   20
         Top             =   2940
         Width           =   3405
      End
      Begin VB.TextBox txtSig_FinMan 
         Height          =   375
         Left            =   2130
         TabIndex        =   19
         Top             =   2535
         Width           =   3405
      End
      Begin VB.TextBox txtSig_DeliveryBy 
         Height          =   375
         Left            =   2130
         TabIndex        =   18
         Top             =   2145
         Width           =   3405
      End
      Begin VB.TextBox txtSig_SalesDispact 
         Height          =   375
         Left            =   2130
         TabIndex        =   17
         Top             =   1740
         Width           =   3405
      End
      Begin VB.TextBox txtSig_SalesApprove 
         Height          =   375
         Left            =   2130
         TabIndex        =   16
         Top             =   1335
         Width           =   3405
      End
      Begin VB.TextBox txtSig_CheckedBy 
         Height          =   375
         Left            =   2130
         TabIndex        =   15
         Top             =   900
         Width           =   3405
      End
      Begin VB.TextBox txtSig_PrepBy 
         Height          =   375
         Left            =   2130
         TabIndex        =   14
         Top             =   480
         Width           =   3405
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "DESIGNATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         TabIndex        =   24
         Top             =   150
         Width           =   3555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "                                 SIGNATORIES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   23
         Top             =   150
         Width           =   5505
      End
      Begin VB.Label CAP6 
         Caption         =   "FinancingManager"
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   2655
         Width           =   1575
      End
      Begin VB.Label CAP5 
         Caption         =   "Delivered By"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   2265
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "GeneralManager"
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   3090
         Width           =   1575
      End
      Begin VB.Label CAP4 
         Caption         =   "Sales Dispatcher"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   1890
         Width           =   1935
      End
      Begin VB.Label CAP3 
         Caption         =   "Sales Manager"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   1455
         Width           =   1575
      End
      Begin VB.Label CAP2 
         Caption         =   "Checked By"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label CAP1 
         Caption         =   "Prepared/Reviewed By"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   600
         Width           =   1905
      End
   End
   Begin VB.PictureBox picSaves 
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
      Height          =   885
      Left            =   7710
      ScaleHeight     =   885
      ScaleWidth      =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   2760
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   690
         MouseIcon       =   "Signatories.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Signatories.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   0
         MouseIcon       =   "Signatories.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Signatories.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3330
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label labid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   1
      Top             =   4650
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmSMIS_Files_Signatories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSignatories                                                     As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub CheckSignatories()
    Dim rsTemp                                                        As ADODB.Recordset


    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='TRANSACTION SLIP' AND MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('TRANSACTION SLIP','SMIS')")
    End If
    
    
        Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='JOB REQUEST FORM' AND MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('JOB REQUEST FORM','SMIS')")
    End If
    
    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='RELEASE ORDER' AND MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO smis_signatories (USEDIN,MAINMODULENAME) VALUES('RELEASE ORDER','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='RECIEVING REPORT' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('RECIEVING REPORT','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='GATE PASS' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('GATE PASS','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='SALES INVOICE' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('SALES INVOICE','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='SALES ORDER' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('SALES ORDER','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='PURCHASE ORDER' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('PURCHASE ORDER','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='DEBIT MEMO' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('DEBIT MEMO','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='CREDIT MEMO' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('CREDIT MEMO','SMIS')")
    End If

    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='DELIVERY REPORT' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('DELIVERY REPORT','SMIS')")
    End If
    
    Set rsTemp = gconDMIS.Execute("select count(*) from smis_signatories where USEDIN ='STOCK TRANSFER' and MAINMODULENAME='SMIS'")
    If rsTemp.Fields(0).Value = 0 Then
        gconDMIS.Execute ("INSERT INTO SMIS_SIGNATORIES (USEDIN,MAINMODULENAME)  VALUES('STOCK TRANSFER','SMIS')")
    End If
End Sub

Sub InitMemVars()
    txtSig_CheckedBy = ""
    txtSig_DeliveryBy = ""
    txtSig_FinMan = ""
    txtSig_GM = ""
    txtSig_PrepBy = ""
    txtSig_SalesApprove = ""
    txtSig_SalesDispact = ""

End Sub

Sub rsRefresh()
    Set rsSignatories = New ADODB.Recordset
    rsSignatories.Open "select * from SMIS_SIGNATORIES order by id DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        labid.Caption = rsSignatories!ID
        txtSig_CheckedBy = Null2String(rsSignatories!CheckedBy)

        txtSig_DeliveryBy = Null2String(rsSignatories!DeliveredBy)
        txtSig_FinMan = Null2String(rsSignatories!FinancingManager)
        txtSig_GM = Null2String(rsSignatories!GeneralManager)
        txtSig_PrepBy = Null2String(rsSignatories!PreparedBy)
        'txtSig_PrepBy = GetSetting("DMIS", "SMIS", "SIGNATORIES", LOGNAME)
        txtSig_SalesApprove = Null2String(rsSignatories!SalesApproved)
        txtSig_SalesDispact = Null2String(rsSignatories!SalesDispatcher)
        '------------------------------------------------------------------------------------------------
        'UPDATED BY: JUN
        'DATE UPDATED: 09142008 12:06PM
        'DESCRIPTION: ADDED DESIGNATION OF SIGNATORIES
        
        txtPreRevDesig = Null2String(rsSignatories!PreparedByDesig)
        txtChkDesig = Null2String(rsSignatories!CheckedByDesig)
        txtSManagerDesig = Null2String(rsSignatories!SalesApprovedDesig)
        txtSDispatchDesig = Null2String(rsSignatories!SalesDispatcherDesig)
        txtDelByDesig = Null2String(rsSignatories!DeliveredByDesig)
        txtFinManDesig = Null2String(rsSignatories!FinancingManagerDesig)
        txtGenManDesig = Null2String(rsSignatories!GeneralManagerDesig)
        
        '-------------------------------------------------------------------------------------------------
        Combo1.Text = Null2String(rsSignatories!USEDIN)


        PreparedBy = Null2String(rsSignatories!PreparedBy)
        ApprovedBy = Null2String(rsSignatories!SalesApproved)
        CheckedBy = Null2String(rsSignatories!CheckedBy)
        SalesDispatcher = Null2String(rsSignatories!SalesDispatcher)
        GeneralManager = Null2String(rsSignatories!GeneralManager)
        DeliveredBy = Null2String(rsSignatories!DeliveredBy)
        FinancingManager = Null2String(rsSignatories!FinancingManager)

    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Private Sub cmdSave_Click()
'    On Error GoTo ErrorCode:
'    Dim SQL                                                           As String
'    SQL = " UPDATE SMIS_Signatories SET "
'    SQL = SQL & " DeliveredBy=" & N2Str2Null(txtSig_DeliveryBy) & " ,"
'    SQL = SQL & " CheckedBy=" & N2Str2Null(txtSig_CheckedBy) & " ,"
'    SQL = SQL & " FinancingManager=" & N2Str2Null(txtSig_FinMan) & " ,"
'    SQL = SQL & " GeneralManager=" & N2Str2Null(txtSig_GM) & " ,"
'    SQL = SQL & " PreparedBy=" & N2Str2Null(txtSig_PrepBy) & " ,"
'    SQL = SQL & " SalesApproved=" & N2Str2Null(txtSig_SalesApprove) & " ,"
'    SQL = SQL & " SalesDispatcher=" & N2Str2Null(txtSig_SalesDispact)
'    SQL = SQL & " Where ID=" & labid
'    LogAudit "E", "SIGNATORIES", Combo1
'    gconDMIS.Execute SQL
'    rsREFRESH
'    MessagePop RecSaveOk, "Signatories Record Updated", "Record Sucessfully Updated", 1000
'    rsSignatories.Find ("ID=" & labid)
'
'    StoreMemVars
'ErrorCode:
'    ShowVBError
'End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:
    Dim SQL                                                           As String
    SQL = " UPDATE SMIS_Signatories SET "
    SQL = SQL & " DeliveredBy=" & N2Str2Null(txtSig_DeliveryBy) & " ,"
    SQL = SQL & " CheckedBy=" & N2Str2Null(txtSig_CheckedBy) & " ,"
    SQL = SQL & " FinancingManager=" & N2Str2Null(txtSig_FinMan) & " ,"
    SQL = SQL & " GeneralManager=" & N2Str2Null(txtSig_GM) & " ,"
    SQL = SQL & " PreparedBy=" & N2Str2Null(txtSig_PrepBy) & " ,"
    SQL = SQL & " SalesApproved=" & N2Str2Null(txtSig_SalesApprove) & " ,"
    SQL = SQL & " SalesDispatcher=" & N2Str2Null(txtSig_SalesDispact) & " ,"
    SQL = SQL & " PreparedByDesig=" & N2Str2Null(txtPreRevDesig) & " ,"
    SQL = SQL & " CheckedByDesig=" & N2Str2Null(txtChkDesig) & " ,"
    SQL = SQL & " SalesApprovedDesig=" & N2Str2Null(txtSManagerDesig) & " ,"
    SQL = SQL & " SalesDispatcherDesig=" & N2Str2Null(txtSDispatchDesig) & " ,"
    SQL = SQL & " GeneralManagerDesig=" & N2Str2Null(txtGenManDesig) & " ,"
    SQL = SQL & " DeliveredByDesig=" & N2Str2Null(txtDelByDesig) & " ,"
    SQL = SQL & " FinancingManagerDesig=" & N2Str2Null(txtFinManDesig)
    SQL = SQL & " Where ID=" & labid
    LogAudit "E", "SIGNATORIES", Combo1
    gconDMIS.Execute SQL
    rsRefresh
    MessagePop RecSaveOk, "Signatories Record Updated", "Record Sucessfully Updated", 1000
    rsSignatories.Find ("ID=" & labid)

    StoreMemVars
ErrorCode:
    ShowVBError
End Sub



Private Sub Combo1_Change()
    Combo1_Click
End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    rsSignatories.MoveFirst
    rsSignatories.Find ("USEDIN='" & Combo1.Text & "'")
    InitMemVars
    StoreMemVars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CheckSignatories
    rsRefresh
    InitMemVars
    Combo_Loadval Combo1, gconDMIS.Execute("Select DISTINCT USEDIN from SMIS_SIGNATORIES ORDER BY USEDIN")
    StoreMemVars
    Screen.MousePointer = 0
End Sub

