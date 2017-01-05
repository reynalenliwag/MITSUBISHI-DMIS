VERSION 5.00
Begin VB.Form frmSMIS_Files_Signatories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signatories"
   ClientHeight    =   4980
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5790
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
   Icon            =   "frmSMIS_Files_Signatories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   5790
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5715
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmSMIS_Files_Signatories.frx":08CA
         Left            =   1860
         List            =   "frmSMIS_Files_Signatories.frx":08CC
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   240
         Width           =   3405
      End
      Begin VB.Label Label1 
         Caption         =   "Forms/Application"
         Height          =   255
         Left            =   150
         TabIndex        =   7
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
      Height          =   3375
      Left            =   30
      TabIndex        =   3
      Top             =   630
      Width           =   5715
      Begin VB.TextBox txtSig_GM 
         Height          =   375
         Left            =   1830
         TabIndex        =   22
         Top             =   2670
         Width           =   3405
      End
      Begin VB.TextBox txtSig_FinMan 
         Height          =   375
         Left            =   1830
         TabIndex        =   21
         Top             =   2270
         Width           =   3405
      End
      Begin VB.TextBox txtSig_DeliveryBy 
         Height          =   375
         Left            =   1830
         TabIndex        =   20
         Top             =   1870
         Width           =   3405
      End
      Begin VB.TextBox txtSig_SalesDispact 
         Height          =   375
         Left            =   1830
         TabIndex        =   19
         Top             =   1470
         Width           =   3405
      End
      Begin VB.TextBox txtSig_SalesApprove 
         Height          =   375
         Left            =   1830
         TabIndex        =   18
         Top             =   1070
         Width           =   3405
      End
      Begin VB.TextBox txtSig_CheckedBy 
         Height          =   375
         Left            =   1830
         TabIndex        =   17
         Top             =   670
         Width           =   3405
      End
      Begin VB.TextBox txtSig_PrepBy 
         Height          =   375
         Left            =   1830
         TabIndex        =   16
         Top             =   270
         Width           =   3405
      End
      Begin VB.Label Label8 
         Caption         =   "FinancingManager"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2325
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Delivered By"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1965
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "GeneralManager"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "SalesDispatcher"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "SalesApproved"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1155
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "CheckedBy"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   735
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "PreparedBy"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   330
         Width           =   1575
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
      Left            =   3600
      ScaleHeight     =   885
      ScaleWidth      =   2760
      TabIndex        =   4
      Top             =   4020
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
         Left            =   1440
         MouseIcon       =   "frmSMIS_Files_Signatories.frx":08CE
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Files_Signatories.frx":0A20
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   750
         MouseIcon       =   "frmSMIS_Files_Signatories.frx":0D86
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Files_Signatories.frx":0ED8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save this Record"
         Top             =   30
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
      Left            =   8160
      TabIndex        =   2
      Top             =   570
      Width           =   195
   End
   Begin VB.Label labid 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   1
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmSMIS_FILES_Signatories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSignatories                                                     As ADODB.Recordset
Dim AddorEdit                                                         As String

Private Sub cmdExit_Click()
    Unload Me
End Sub


'Upating Code       : AXP-0707200712:20
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
    SQL = SQL & " SalesDispatcher=" & N2Str2Null(txtSig_SalesDispact)
    SQL = SQL & " Where ID=" & labid
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
    rsSignatories.MoveFirst
    rsSignatories.Find ("USEDIN='" & Combo1.Text & "'")
    initMemvars
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
    rsRefresh
    initMemvars
    Combo_Loadval Combo1, gconDMIS.Execute("Select DISTINCT upper(USEDIN) from SMIS_SIGNATORIES")
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
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
    rsSignatories.Open "select * from SMIS_SIGNATORIES order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        labid.Caption = rsSignatories!ID
        txtSig_CheckedBy = Null2String(rsSignatories!CheckedBy)
        txtSig_DeliveryBy = Null2String(rsSignatories!DeliveredBy)
        txtSig_FinMan = Null2String(rsSignatories!FinancingManager)
        txtSig_GM = Null2String(rsSignatories!GeneralManager)
        txtSig_PrepBy = Null2String(rsSignatories!PreparedBy)
        txtSig_SalesApprove = Null2String(rsSignatories!SalesApproved)
        txtSig_SalesDispact = Null2String(rsSignatories!SalesDispatcher)
        Combo1.Text = Null2String(rsSignatories!USEDIN)
    Else
        ShowNoRecord

    End If
End Sub

