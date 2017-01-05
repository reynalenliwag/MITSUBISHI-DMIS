VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_Confilct 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConflictingRecord.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4725
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
      _Version        =   655364
      _ExtentX        =   7858
      _ExtentY        =   8334
      _StockProps     =   64
      BorderStyle     =   2
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6825
      Left            =   4560
      ScaleHeight     =   6795
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   0
      Width           =   2745
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3315
         Left            =   0
         ScaleHeight     =   3315
         ScaleWidth      =   4050
         TabIndex        =   6
         Top             =   3450
         Width           =   4050
         Begin VB.TextBox txtCusName 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   525
            HideSelection   =   0   'False
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCusEmail 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            HideSelection   =   0   'False
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1980
         End
         Begin VB.TextBox txtCusAdd 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1110
            HideSelection   =   0   'False
            Left            =   15
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2190
            Width           =   2655
         End
         Begin VB.TextBox txtCusPhone 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            HideSelection   =   0   'False
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1650
            Width           =   1980
         End
         Begin VB.TextBox lblNotes 
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   30
            Width           =   2655
         End
         Begin VB.Label lblCustAddress 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   15
            TabIndex        =   14
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label lblCustPhone 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Phone:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   30
            TabIndex        =   13
            Top             =   1650
            Width           =   675
         End
         Begin VB.Label lblCustEmail 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Email :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   30
            TabIndex        =   12
            Top             =   1380
            Width           =   675
         End
      End
      Begin VB.Label lblLogLetters 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   35
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblLogCalls 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   34
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label lblLogQuote 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   33
         ToolTipText     =   " Last Quotation Send "
         Top             =   345
         Width           =   1695
      End
      Begin VB.Label lblAppt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Appointment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   30
         TabIndex        =   32
         Top             =   915
         Width           =   915
      End
      Begin VB.Label lblQ 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Quotation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   31
         Top             =   345
         Width           =   930
      End
      Begin VB.Label lblCalls 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Calls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   30
         Top             =   1755
         Width           =   930
      End
      Begin VB.Label lblLetters 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Letters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   29
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label lblLogTestDrive 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   28
         ToolTipText     =   " Test Drive Schedules On and Day Elasped"
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label lblTest 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Test Drive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   27
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lblLogAppointment 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   960
         TabIndex        =   26
         ToolTipText     =   "Last sales appointment made on and days elasped"
         Top             =   915
         Width           =   1695
      End
      Begin VB.Label lblVisits 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Visits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   25
         Top             =   1470
         Width           =   930
      End
      Begin VB.Label lblLogVisits 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   24
         Top             =   1470
         Width           =   1695
      End
      Begin VB.Label lblLoan 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Loan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   23
         Top             =   2610
         Width           =   930
      End
      Begin VB.Label lblLogLoan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   22
         Top             =   2610
         Width           =   1695
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   21
         Top             =   2895
         Width           =   2640
      End
      Begin VB.Label lblSalesOrder 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Sales order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   20
         Top             =   1185
         Width           =   930
      End
      Begin VB.Label lblLogSalesOrder 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   19
         Top             =   1185
         Width           =   1695
      End
      Begin VB.Label lblLogEmail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   960
         TabIndex        =   18
         Top             =   2325
         Width           =   1695
      End
      Begin VB.Label lblEmails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   17
         Top             =   2325
         Width           =   930
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   0
         Y1              =   300
         Y2              =   3465
      End
      Begin VB.Label lblAgeing 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   15
         TabIndex        =   16
         Top             =   3180
         Width           =   2640
      End
      Begin XtremeShortcutBar.ShortcutCaption captionInformation 
         Height          =   315
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   2715
         _Version        =   655364
         _ExtentX        =   4789
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Profile @"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   3810
      MouseIcon       =   "ConflictingRecord.frx":01CA
      MousePointer    =   99  'Custom
      Picture         =   "ConflictingRecord.frx":031C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   6000
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   4875
      Left            =   9660
      TabIndex        =   2
      Text            =   "lblError"
      Top             =   -60
      Width           =   3015
   End
   Begin VB.TextBox lblError 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ConflictingRecord.frx":065A
      Top             =   60
      Width           =   4455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&View"
      Height          =   735
      Left            =   3120
      MouseIcon       =   "ConflictingRecord.frx":0663
      MousePointer    =   99  'Custom
      Picture         =   "ConflictingRecord.frx":07B5
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Select"
      Top             =   6000
      Width           =   705
   End
End
Attribute VB_Name = "frmSMIS_Trans_Confilct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConflictingKey                                                    As String
Dim ConflictingKeyWords                                               As String
Dim CUSCDE                                                            As String
Public Event SelectionMade(oRs As ADODB.Recordset)

Sub ShowStatus(PROSPECTID)
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim xstatus                                                       As String
    Dim CustomerCode                                                  As String
    Dim ProspType                                                     As String
    Dim EXIST_LOAN                                                    As Boolean
    Dim Exist_SO                                                      As Boolean
    Set TEMPRS = gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)
    lblLogQuote = "": lblLogEmail = "": lblLogAppointment = "": lblLogTestDrive = "": lblLogCalls = ""
    lblLogLetters = "": lblLetters = "": lblLogLoan = "": lblLogVisits = "": lblLogSalesOrder = "": lblSTATUS = ""
    txtCusAdd = "": txtCusEmail = "": txtCusName = "": txtCusPhone = "": lblNotes = "": captionInformation.Caption = ""
    CustomerCode = "": captionInformation.Caption = ""
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        lblLogQuote = "x": lblLogQuote.ForeColor = &HC0&
        lblLogEmail = "x": lblLogEmail.ForeColor = &HC0&
        lblLogAppointment = "x": lblLogAppointment.ForeColor = &HC0&
        lblLogTestDrive = "x": lblLogTestDrive.ForeColor = &HC0&
        lblLogCalls = "x": lblLogCalls.ForeColor = &HC0&
        lblLogLetters = "x": lblLogLetters.ForeColor = &HC0&
        lblLogLoan = "x": lblLogLoan.ForeColor = &HC0&
        lblSTATUS = "x": lblSTATUS.BackColor = vbWhite
        lblLogVisits = "x": lblLogVisits.ForeColor = &HC0&
        lblLogSalesOrder.Caption = "x": lblLogSalesOrder.ForeColor = &HC0&
        lblNotes.Text = ""
        ProspType = Null2String(TEMPRS!ProspectType)
        lblNotes.Text = Null2String(TEMPRS!Notes)

        If IsNull(TEMPRS!LogQuote) = False Then
            lblLogQuote = Chr(187) & FormatDateTime(TEMPRS!LogQuote, vbShortDate): lblLogQuote.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogEmail) = False Then
            lblLogEmail = Chr(187) & FormatDateTime(TEMPRS!LogEmail, vbShortDate): lblLogEmail.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogAppointment) = False Then
            lblLogAppointment = Chr(187) & FormatDateTime(TEMPRS!LogAppointment, vbShortDate): lblLogAppointment.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogTestDrive) = False Then
            lblLogTestDrive = Chr(187) & FormatDateTime(TEMPRS!LogTestDrive, vbShortDate) & "(" & DateDiff("d", TEMPRS!LogTestDrive, LOGDATE) & ")": lblLogTestDrive.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogCall) = False Then
            lblLogCalls = Chr(187) & FormatDateTime(TEMPRS!LogCall, vbShortDate): lblLogCalls.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogLetter) = False Then
            lblLogLetters = Chr(187) & FormatDateTime(TEMPRS!LogLetter, vbShortDate): lblLogLetters.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogApplication) = False Then
            EXIST_LOAN = True
            lblLogLoan = Chr(187) & FormatDateTime(TEMPRS!LogApplication, vbShortDate): lblLogLoan.ForeColor = &H8000&
        End If
        If IsNull(TEMPRS!LogVisit) = False Then
            lblLogVisits = Chr(187) & FormatDateTime(TEMPRS!LogVisit, vbShortDate): lblLogVisits.ForeColor = &H8000&
        End If
        lblAgeing = "Aging: " & DateDiff("d", TEMPRS!loginitialinquiry, LOGDATE) & " Days"
        If IsNull(TEMPRS!LOGSO) = False Then
            Exist_SO = True
            lblLogSalesOrder = Chr(187) & FormatDateTime(TEMPRS!LOGSO, vbShortDate): lblLogSalesOrder.ForeColor = &H8000&
        End If
        xstatus = Null2String(TEMPRS!STATUS)
        If xstatus = "O" Then
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        ElseIf xstatus = "C" Then
            lblSTATUS = "CLOSED": lblSTATUS.BackColor = &H40C0&
        ElseIf xstatus = "I" Then
            lblSTATUS = "INACTIVE": lblSTATUS.BackColor = &HC0C0C0
        Else
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        End If
        If IsNull(TEMPRS!CUSCDE) = False Then
            CustomerCode = TEMPRS!CUSCDE: ShowCustomerInfo CustomerCode
        End If
    End If
    Set TEMPRS = Nothing
End Sub

Friend Sub Conflict(xConflictingKey As String, xConflictingKeyWords As String, Optional ByVal xCuscde As String)
    ConflictingKey = xConflictingKey
    ConflictingKeyWords = xConflictingKeyWords
    CUSCDE = xCuscde
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub lvGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorCode:

    If KeyCode = 13 Then
        Call lvGrid_RowDblClick(lvGrid.Rows(lvGrid.SelectedRows.Row(0).Index), Nothing)
    End If
    If KeyCode = vbKeyF3 Then
        Call frmSMIS_Mis_Filter.ConfigGrid(lvGrid, 0)
        frmSMIS_Mis_Filter.Show 1
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then
        Exit Sub
    Else
        Dim TEMPRS                                                    As ADODB.Recordset
        Set TEMPRS = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & Row.Record(4).Value)
        If Not TEMPRS Is Nothing Then
            RaiseEvent SelectionMade(TEMPRS)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    lvGrid_KeyDown 13, 0
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    ReportControlAddColumnHeader lvGrid, "PROSPECTNAME, EMAIL, MOBILE, PHONE"
    ReportControlPaintManager lvGrid
    ResizeColumnHeader lvGrid, "40, 20,20,20"

    If Len(Trim(CUSCDE)) = 0 Then
        flex_FillReportView gconDMIS.Execute("SELECT ACCTNAME, EMAIL, MOBILE, TELEPHONE,  PROSPECTID from CRIS_PROSPECTS Where " & ConflictingKey & "=" & N2Str2Null(ConflictingKeyWords)), lvGrid, False
    Else
        flex_FillReportView gconDMIS.Execute("SELECT ACCTNAME, EMAIL, MOBILE, TELEPHONE,  PROSPECTID from CRIS_PROSPECTS Where PROSPECTID<>" & N2Str2Null(CUSCDE) & " AND " & ConflictingKey & "=" & N2Str2Null(ConflictingKeyWords)), lvGrid, False
    End If
End Sub

Private Sub lvGrid_SelectionChanged()

    On Error GoTo ErrorCode

    cmdok.Enabled = True
    ShowStatus lvGrid.SelectedRows(0).Record(4).Value


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub ShowCustomerInfo(xxxcode)
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(xxxcode))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        If TEMPRS.Fields("CUSTYPE") = "P" Then
            txtCusAdd = Replace(Null2String(TEMPRS("CUSTOMERADD")), Chr(10), "")
            txtCusEmail = Null2String(TEMPRS("EMAIL"))
            txtCusName = Null2String(TEMPRS("LASTNAME")) & IIf(IsNull(TEMPRS("LASTNAME")), "", ",") & Null2String(TEMPRS("FirstName")) & IIf(IsNull(TEMPRS("MIDDLEINITIAL")), "", ".") & Null2String(TEMPRS("MIDDLEINITIAL"))
            txtCusPhone = Null2String(TEMPRS("HOMEPHONE")) & " /" & Null2String(TEMPRS("TELEPHONENO"))
            captionInformation.Caption = Null2String(TEMPRS("ACCTNAME"))
            lblSTATUS = " LastUpdated:" & Null2String(TEMPRS("LASTUPDATE"))
        Else
            txtCusAdd = Replace(Null2String(TEMPRS("COMPANYADD")), Chr(10), "")
            txtCusEmail = Null2String(TEMPRS("EMAIL"))
            txtCusName = Null2String(TEMPRS("CUSCOMP"))
            txtCusPhone = Null2String(TEMPRS("HOMEPHONE")) & " /" & Null2String(TEMPRS("TELEPHONENO"))
            captionInformation.Caption = Null2String(TEMPRS("ACCTNAME"))
            lblSTATUS = " LastUpdated:" & Null2String(TEMPRS("LASTUPDATE"))
        End If
    End If
    Set TEMPRS = Nothing
End Sub

