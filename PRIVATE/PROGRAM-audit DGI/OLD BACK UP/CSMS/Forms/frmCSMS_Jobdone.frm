VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_Jobdone 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Job Done"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Jobdone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   30
      ScaleHeight     =   2565
      ScaleWidth      =   6765
      TabIndex        =   12
      Top             =   2460
      Width           =   6795
      Begin VB.TextBox txtstdrate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4965
         TabIndex        =   1
         Top             =   390
         Width           =   1755
      End
      Begin VB.TextBox txtflatrate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         TabIndex        =   0
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtjobdesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   780
         Width           =   5415
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Left            =   -15
         TabIndex        =   28
         Top             =   0
         Width           =   6765
         _Version        =   655364
         _ExtentX        =   11933
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   "Final Findings"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Details:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   300
         TabIndex        =   15
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Std. TIme:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4065
         TabIndex        =   14
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat Rate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   510
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6090
      MouseIcon       =   "frmCSMS_Jobdone.frx":0DCA
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Jobdone.frx":0F1C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   5085
      Width           =   735
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   7380
      TabIndex        =   6
      Top             =   2700
      Width           =   2415
      Begin VB.TextBox txtJobDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   7500
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label55 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   7530
         TabIndex        =   9
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6990
         TabIndex        =   8
         Top             =   780
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5340
      MouseIcon       =   "frmCSMS_Jobdone.frx":125A
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Jobdone.frx":13AC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save Entry"
      Top             =   5085
      Width           =   765
   End
   Begin VB.Frame Frame8 
      Caption         =   "Job Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   7350
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   30
      ScaleHeight     =   2385
      ScaleWidth      =   6765
      TabIndex        =   16
      Top             =   30
      Width           =   6795
      Begin VB.TextBox lbljobdesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1770
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1080
         Width           =   4905
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   300
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   6765
         _Version        =   655364
         _ExtentX        =   11933
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   "Job Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label LABITEMNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   27
         Top             =   1710
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order no:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   345
         TabIndex        =   26
         Top             =   450
         Width           =   1350
      End
      Begin VB.Label lblJobCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1770
         TabIndex        =   25
         Top             =   750
         Width           =   4905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   870
         TabIndex        =   24
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job  Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   315
         TabIndex        =   23
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label lblro 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   22
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flat Rate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   960
         TabIndex        =   21
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Std. Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4080
         TabIndex        =   20
         Top             =   1950
         Width           =   825
      End
      Begin VB.Label lblflatrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1770
         TabIndex        =   19
         Top             =   1920
         Width           =   1725
      End
      Begin VB.Label lblstdrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4980
         TabIndex        =   18
         Top             =   1920
         Width           =   1725
      End
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  'Transparent
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
      Left            =   1350
      TabIndex        =   11
      Top             =   5130
      Width           =   3855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMATION :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   45
      TabIndex        =   10
      Top             =   5115
      Width           =   1200
   End
End
Attribute VB_Name = "frmCSMS_Jobdone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim JOBDETAMT                                          As Double
Dim JOBTAXVAL                                          As Double
Dim AUDIT_SQL                                          As String

Sub updateJob()
    'JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
End Sub

Sub ComputeMe()
    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim VLastUpdateTime                                 As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim VLastUpdate                                     As String
    Dim Vusercode                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim xApptNo                                         As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim xFLATRATE                                       As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As Double
    Dim TheDone                                         As String
    Dim JOBDETCOST                                      As Double
    Dim VROTYPE                                         As String
    Dim vJobType                                        As String

    If txtflatrate.Text = "" Then
        MsgBox "Please Input FlatRate..", vbExclamation, "Information"
        txtflatrate.SetFocus
        Exit Sub
    End If

    If txtstdrate.Text = "" Then
        MsgBox "Please input StdRate..", vbExclamation, "Information"
        txtstdrate.SetFocus
        Exit Sub
    End If

    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
    JOBREP_OR = N2Str2Null(lblro.Caption)
    JOBDET_HRS = NumericVal(txtstdrate.Text)
    xFLATRATE = NumericVal(txtflatrate.Text)
    JOBDETPRC = NumericVal(txtflatrate.Text) * JOBDET_HRS
    JOBTAXRATE = (VAT_RATE / 100)
    JOBDISCRATE = NumericVal(txtJobDiscount.Text) / 100
    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE

    'COMMENT BY  : MJP 10162009 1029 AM
    'DESCRIPTION : DOUBLE VAT
        'JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'COMMENT BY  : MJP 10162009 1029 AM
    'UPDATE BY   : MJP 10162009 1029 AM
        JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    'UPDATE BY   : MJP 10162009 1029 AM
    
    If MsgBox("Are you sure you want to continue?", vbQuestion + vbYesNo) = vbYes Then
        If lblJobCode.Caption = "SRLABOR" Or lblJobCode.Caption = "SRPARTS" Or lblJobCode.Caption = "SRMATERIALS" Then
            SQL_STATEMENT = "Update CSMS_ro_det set " & _
                " detail = '" & RTrim(LTrim(Replace(txtjobdesc, vbCrLf, " "))) & _
                "' where rep_or = " & JOBREP_OR & _
                " and detcde = '" & lblJobCode & _
                "' and livil = '1' " & _
                " AND LINE_NO = '" & LABITEMNO & "'"
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(JOBREP_OR), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & lblJobCode & "JOB DONE", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            
            lblflatrate.Caption = txtflatrate.Text
            lblstdrate.Caption = txtstdrate.Text

            Call ShowSuccessFullyUpdated
            Call cmdCancel_Click
        Else
            SQL_STATEMENT = "Update CSMS_ro_det set " & _
                " det_hrs = " & JOBDET_HRS & _
                ", flatrate = " & xFLATRATE & _
                ", det_amt = " & JOBDETPRC & _
                ", detamt = " & JOBDETAMT & _
                ", taxval = " & JOBTAXVAL & _
                ", detprc = " & JOBDETPRC & _
                ", detail = '" & RTrim(LTrim(txtjobdesc.Text)) & _
                "' where rep_or = " & JOBREP_OR & _
                " and detcde = '" & lblJobCode & _
                "' and livil = '1' " & _
                " AND LINE_NO = '" & LABITEMNO & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(JOBREP_OR), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & lblJobCode & "JOB DONE", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            
            lblflatrate.Caption = txtflatrate.Text
            lblstdrate.Caption = txtstdrate.Text

            Call ShowSuccessFullyUpdated
            
            Call cmdCancel_Click
            Call frmCSMS_ServiceCounter.Click_ScheduleGrid
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If NumericVal(txtstdrate) < 0.1 Then
        On Error Resume Next
        MsgBox "Std. Rate Cannot be Zero/Blank or Less than 6 minute", vbExclamation, "CSMS"
        txtstdrate.SetFocus
        Exit Sub
    End If

    Call ComputeMe
    'cmdOk.Visible = False
    'cmdCancel.Caption = "Close"
End Sub

Private Sub Form_Load()
    'CenterMe frmMain, Me, 0
End Sub

Private Sub txtflatrate_GotFocus()
    lblinfo.Caption = "Input the final Flatrate"
    txtflatrate.BackColor = &HC0FFFF
    txtstdrate.BackColor = vbWhite
    txtjobdesc.BackColor = vbWhite
End Sub

Private Sub txtflatrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtjobdesc_GotFocus()
    lblinfo.Caption = "Input the final Description"
    txtflatrate.BackColor = vbWhite
    txtstdrate.BackColor = vbWhite
    txtjobdesc.BackColor = &HC0FFFF
End Sub

Private Sub txtjobdesc_LostFocus()
    txtjobdesc.BackColor = vbWhite
End Sub

Private Sub txtstdrate_GotFocus()
    lblinfo.Caption = "Input the final std Rate"
    txtflatrate.BackColor = vbWhite
    txtstdrate.BackColor = &HC0FFFF
    txtjobdesc.BackColor = vbWhite
End Sub

Private Sub txtstdrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub
