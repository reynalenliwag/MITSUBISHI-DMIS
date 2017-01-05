VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISProcess_UpdateAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Adjustment File"
   ClientHeight    =   1545
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   5745
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "UpdateAdjustment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5745
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
      Left            =   4860
      MouseIcon       =   "UpdateAdjustment.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAdjustment.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit Window"
      Top             =   690
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
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
      Left            =   4140
      MouseIcon       =   "UpdateAdjustment.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateAdjustment.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Process Adjustment File"
      Top             =   690
      Width           =   735
   End
   Begin MSMask.MaskEdBox txtTrandate 
      Height          =   345
      Left            =   2460
      TabIndex        =   0
      ToolTipText     =   "Input valid transaction date"
      Top             =   1110
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mmm-yy"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   30
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   3
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   4
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Update progress"
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateAdjustment.frx":0BAF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateAdjustment.frx":0BCB
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   5
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   6
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "UpdateAdjustment.frx":0BE7
         End
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Transaction Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1140
      Width           =   2385
   End
End
Attribute VB_Name = "frmPMISProcess_UpdateAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdjust                                           As ADODB.Recordset
Dim vAdd                                               As Double
Dim vMinus                                             As Double
Dim i                                                  As Integer
Dim vTrandate, vPARTNO                                 As String
Dim vID                                                As Integer
Dim VStatus                                            As String
Dim rsPartsAdjust                                      As ADODB.Recordset
Dim rsPmasMAC                                          As ADODB.Recordset
Dim AdjustQty_Add, AdjustQty_Minus                     As Integer
Dim vMAC                                               As Double
Dim iqty                                               As Integer
Dim XTYPE                                              As String
Dim RLMAC                                              As Double

Sub UpdateAdjustment()

    rsAdjust.MoveFirst
    Screen.MousePointer = 11
    Me.Caption = "Updating Adjustment to Transactions Master File"
    DoEvents
    i = 0
start:
    Do While Not rsAdjust.EOF
        vID = rsAdjust!ID
        vTrandate = N2Date2Null(txttrandate.Text)
        vPARTNO = N2Str2Null(rsAdjust!PARTNO)
        vMinus = N2Str2Zero(rsAdjust!minus)
        vAdd = N2Str2Zero(rsAdjust!Add)
        XTYPE = Null2String(rsAdjust!Type)
        VStatus = "'N'"

        'updating code:     JAA - 09062008    -  Get MAC from Stockmas
        Set rsPmasMAC = New ADODB.Recordset
        Set rsPmasMAC = gconDMIS.Execute("SELECT MAC,onhand FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(vPARTNO))
        If Not rsPmasMAC.EOF And Not rsPmasMAC.BOF Then
            vMAC = N2Str2Zero(rsPmasMAC!MAC)
            iqty = NumericVal(rsPmasMAC!ONHAND) - vMinus
        'updated BY: IEBV_04292011
        'description: To Avoid negative onhand
        '-------------------------------------------------------------------------------------------
            If vMinus > NumericVal(rsPmasMAC!ONHAND) Then
                MsgBox "Cannot Post Adjustment on ( " & vPARTNO & "). This will result to negative onhand.", vbInformation + vbOKOnly
                DoEvents
                i = i + 1
                progCPB.Value = (i / rsAdjust.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsAdjust.MoveNext
                GoTo start
            End If
        '-------------------------------------------------------------------------------------------
        Else
            vMAC = 0
        End If

        If vAdd <> 0 Then
            'updating code:     JAA - 09062008    -  Add MAC and Tranucost in Inserting Adjustment in Tdaytran
            gconDMIS.Execute "Insert into PMIS_TdayTran " & _
                "(TYPE, MAC, TRANUCOST, trandate, trantype, STOCK_ORD, STOCK_SUP, status, tranqty, tranno, itemno, in_out, usercode)" & _
                " values(" & N2Str2Null(rsAdjust!Type) & _
                ", " & vMAC & _
                ", " & N2Str2Null(rsAdjust!COST) & _
                ", " & vTrandate & _
                ", 'ADJ' " & _
                ", " & vPARTNO & _
                ", " & vPARTNO & _
                ", 'P' " & _
                ", " & vAdd & _
                ", '111111' " & _
                ", '1111' " & _
                ", 'I' " & _
                ", " & N2Str2Null(rsAdjust!USERCODE) & ")"

            'updating code:     JAA - 09062008    -  Update the Stock Master File whenever User process the Adjustment
            Set rsPartsAdjust = New ADODB.Recordset
            Set rsPartsAdjust = gconDMIS.Execute("select STOCKNO,Onhand,trecqty,receipts from PMIS_STOCKMAS where TYPE = " & N2Str2Null(rsAdjust!Type) & " AND STOCKNO = " & N2Str2Null(vPARTNO))
            AdjustQty_Add = N2Str2Zero(rsPartsAdjust!ONHAND) + vAdd
            If Not rsPartsAdjust.EOF And Not rsPartsAdjust.BOF Then
                'updating code:     JAA - 09092008    -  Update the trecqty and receipts
                gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                    " onhand = " & AdjustQty_Add & _
                    ", trecqty = " & vAdd + N2Str2Zero(rsPartsAdjust!TRECQTY) & _
                    ", receipts = " & vAdd + N2Str2Zero(rsPartsAdjust!RECEIPTS) & _
                    " where TYPE = " & N2Str2Null(rsAdjust!Type) & _
                    " AND STOCKNO = " & N2Str2Null(vPARTNO)
            End If
        Else
            'updating code:     JAA - 09062008    -  Add MAC and Tranucost in Inserting Adjustment in Tdaytran
            gconDMIS.Execute "Insert into PMIS_TdayTran " & _
                "(TYPE, MAC, TRANUCOST, trandate, trantype, STOCK_ORD, STOCK_SUP, status, tranqty, tranno, itemno, in_out, usercode)" & _
                " values(" & N2Str2Null(rsAdjust!Type) & _
                ", " & vMAC & _
                ", " & N2Str2Null(rsAdjust!COST) & _
                ", " & vTrandate & _
                ", 'ADJ' " & _
                ", " & vPARTNO & _
                ", " & vPARTNO & _
                ", 'P' " & _
                ", " & vMinus & _
                ", '000000' " & _
                ", '0000' " & _
                ", 'O' " & _
                ", " & N2Str2Null(rsAdjust!USERCODE) & ")"

            'updating code:     JAA - 09062008    -  Update the Stock Master File whenever User process the Adjustment
            Set rsPartsAdjust = New ADODB.Recordset
            Set rsPartsAdjust = gconDMIS.Execute("select STOCKNO,Onhand,tissqty,issuances from PMIS_STOCKMAS where TYPE = " & N2Str2Null(rsAdjust!Type) & " AND STOCKNO = " & N2Str2Null(vPARTNO))
            AdjustQty_Minus = N2Str2Zero(rsPartsAdjust!ONHAND) - vMinus
            
            If Not rsPartsAdjust.EOF And Not rsPartsAdjust.BOF Then
                gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                    " onhand = " & AdjustQty_Minus & ", " & _
                    " tissqty = " & N2Str2Zero(rsPartsAdjust!TISSQTY) + vMinus & ", " & _
                    " issuances = " & N2Str2Zero(rsPartsAdjust!ISSUANCES) + vMinus & _
                    " where TYPE = " & N2Str2Null(rsAdjust!Type) & _
                    " AND STOCKNO = " & N2Str2Null(vPARTNO)
            End If

        End If
        gconDMIS.Execute "update PMIS_Adjust set status = 'P' where id = " & vID
        DoEvents
        i = i + 1
        progCPB.Value = (i / rsAdjust.RecordCount) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
        rsAdjust.MoveNext
    Loop
    
    labProcessing.Caption = ""
    DoEvents
    MsgBox "Adjustment Complete.", vbInformation, "PMIS"
    Screen.MousePointer = 0
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "UPDATE ADJUSTMENT FILE") = False Then Exit Sub
    Dim rsmontendna                                     As ADODB.Recordset
    
    On Error GoTo ErrorCode:

    If IsDate(txttrandate.Text) = False Then
        MsgSpeechBox "Please Input Valid Transaction Date!"
        Exit Sub
    End If
    
    Set rsmontendna = gconDMIS.Execute(" select DATE_GEN from PMIS_StkStat where DATE_GEN > '" & CDate(txttrandate.Text) & "' ")
    If Not (rsmontendna.EOF And rsmontendna.BOF) Then
        MsgBox "Monthend has been process in this date, Cannot process update.", vbInformation
        Exit Sub
    End If
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "select * from PMIS_Adjust where status = 'N' AND LASTUPDATE = '" & CDate(txttrandate.Text) & "' order by PARTNO asc", gconDMIS
    If rsAdjust.EOF And rsAdjust.BOF Then
        MsgSpeechBox "Error: Adjustment File is Empty or Adjustments had been Posted already!"
        Exit Sub
    Else
        txttrandate.Enabled = False
        cmdCheck.Enabled = False
        cmdExit.Enabled = False
        DoEvents
        Call UpdateAdjustment

        NEW_LogAudit "R", "UPDATE ADJUSTMENT FILE", "", "", "", txttrandate, "", ""

        cmdExit.Enabled = True
        DoEvents
    End If

    Exit Sub
ErrorCode:
    ShowVBError
    MsgBox Error
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]"   '"." & App.Revision & "]"
    txttrandate.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISProcess_UpdateAdjustment = Nothing
    UnloadForm Me
End Sub

Private Sub txtTrandate_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
