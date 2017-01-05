VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_Jobdone_HPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Job Done"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Jobdone_HPI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
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
      Left            =   9390
      TabIndex        =   27
      Top             =   210
      Width           =   2175
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
      Left            =   9270
      TabIndex        =   22
      Top             =   2670
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
         TabIndex        =   23
         Text            =   "0"
         Top             =   660
         Width           =   495
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
         TabIndex        =   25
         Top             =   780
         Width           =   1095
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
         TabIndex        =   24
         Top             =   330
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   7860
      MouseIcon       =   "frmCSMS_Jobdone_HPI.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Jobdone_HPI.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cancel"
      Top             =   4920
      Width           =   735
   End
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
      Height          =   2505
      Left            =   30
      ScaleHeight     =   2475
      ScaleWidth      =   8535
      TabIndex        =   12
      Top             =   2400
      Width           =   8565
      Begin VB.TextBox txtjobdesc 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   1305
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   750
         Width           =   7125
      End
      Begin VB.TextBox txtflatrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1305
         TabIndex        =   15
         Top             =   390
         Width           =   1665
      End
      Begin VB.TextBox txtstdrate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3945
         TabIndex        =   14
         Top             =   390
         Width           =   1755
      End
      Begin VB.TextBox txtJOBCOST 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6690
         TabIndex        =   13
         Top             =   390
         Width           =   1755
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   8625
         _Version        =   655364
         _ExtentX        =   15214
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
         TabIndex        =   20
         Top             =   450
         Width           =   735
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
         Left            =   3075
         TabIndex        =   19
         Top             =   450
         Width           =   825
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
         TabIndex        =   18
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Cost:"
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
         Left            =   5895
         TabIndex        =   17
         Top             =   420
         Width           =   780
      End
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
      ScaleWidth      =   8535
      TabIndex        =   0
      Top             =   0
      Width           =   8565
      Begin VB.TextBox lbljobdesc 
         BackColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   1290
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1080
         Width           =   7155
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   300
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   8565
         _Version        =   655364
         _ExtentX        =   15108
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
      Begin VB.Label lblstdrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Top             =   1920
         Width           =   1725
      End
      Begin VB.Label lblflatrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1290
         TabIndex        =   10
         Top             =   1920
         Width           =   1665
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
         Left            =   3090
         TabIndex        =   9
         Top             =   1980
         Width           =   825
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
         Left            =   525
         TabIndex        =   8
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label lblro 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1290
         TabIndex        =   7
         Top             =   390
         Width           =   1725
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
         Left            =   435
         TabIndex        =   6
         Top             =   780
         Width           =   825
      End
      Begin VB.Label lblJobCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1290
         TabIndex        =   5
         Top             =   750
         Width           =   7155
      End
      Begin VB.Label lblSTIME 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Cost:"
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
         Left            =   5880
         TabIndex        =   4
         Top             =   1950
         Width           =   780
      End
      Begin VB.Label lblJobCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6690
         TabIndex        =   3
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Job Description:"
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
         Height          =   450
         Left            =   225
         TabIndex        =   2
         Top             =   1140
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   735
      Left            =   7110
      MouseIcon       =   "frmCSMS_Jobdone_HPI.frx":1512
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_Jobdone_HPI.frx":1664
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Save Entry"
      Top             =   4920
      Width           =   765
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
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   30
      TabIndex        =   30
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  'Transparent
      Height          =   765
      Left            =   1380
      TabIndex        =   29
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label LABITEMNO 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "00"
      Height          =   225
      Left            =   360
      TabIndex        =   28
      Top             =   -60
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmCSMS_Jobdone_HPI"
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
    'UPDATED BY: JUN 03-18-2009
    Dim vJOBCOST                                       As Double
    'UPDATED BY: JUN -------------------------
    
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
    
    'UPDATED BY: JUN 03-18-2009
        vJOBCOST = NumericVal(txtJOBCOST.Text)
    'UPDATED BY: JUN -------------------------
    JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)

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
        Else
            SQL_STATEMENT = "Update CSMS_ro_det set det_hrs = " & JOBDET_HRS & _
                ", flatrate = " & xFLATRATE & _
                ", det_amt = " & JOBDETPRC & _
                ", detamt = " & JOBDETAMT & _
                ", taxval = " & JOBTAXVAL & _
                ", detprc = " & JOBDETPRC & _
                ", DETCOST = " & vJOBCOST & _
                ", detail = '" & RTrim(LTrim(Replace(txtjobdesc, vbCrLf, " "))) & _
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
            
            'UPDATED BY: JUN 03-18-2009
                lblJobCost.Caption = txtJOBCOST.Text
            'UPDATED BY: JUN -------------------------
        End If
        
        Call ShowSuccessFullyUpdated
        Call cmdCancel_Click
        Call frmCSMS_ServiceCounter.Click_ScheduleGrid
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
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 0)
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

Private Sub txtJOBCOST_KeyPress(KeyAscii As Integer)
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
