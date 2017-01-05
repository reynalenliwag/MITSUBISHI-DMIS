VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSQUESTION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Print Type"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "frmCSMSQUESTION.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   90
      ScaleHeight     =   1845
      ScaleWidth      =   3015
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3045
      Begin VB.CommandButton cmdEXT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   345
      End
      Begin VB.TextBox txtRO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         MaxLength       =   10
         TabIndex        =   9
         Top             =   780
         Width           =   2685
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter RO no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   10
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdYES2 
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   1590
      Width           =   945
   End
   Begin VB.CommandButton cmdYES1 
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   540
      Width           =   945
   End
   Begin Crystal.CrystalReport rptEstimate 
      Left            =   2010
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Estimate Print Out"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "RO NO"
      Height          =   195
      Index           =   4
      Left            =   3840
      TabIndex        =   17
      Top             =   2130
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "estimate no"
      Height          =   195
      Index           =   3
      Left            =   3540
      TabIndex        =   16
      Top             =   1620
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "discount"
      Height          =   195
      Index           =   2
      Left            =   3750
      TabIndex        =   15
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "date"
      Height          =   195
      Index           =   1
      Left            =   3990
      TabIndex        =   14
      Top             =   750
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      Height          =   195
      Index           =   0
      Left            =   4140
      TabIndex        =   13
      Top             =   330
      Width           =   165
   End
   Begin VB.Label lblRO 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      Height          =   315
      Left            =   4500
      TabIndex        =   11
      Top             =   2010
      Width           =   1815
   End
   Begin VB.Label lblESTI 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      Height          =   315
      Left            =   4500
      TabIndex        =   7
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Label lblDATE 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      Height          =   315
      Left            =   4500
      TabIndex        =   6
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblDISCOUNT 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      Height          =   315
      Left            =   4500
      TabIndex        =   5
      Top             =   1110
      Width           =   1815
   End
   Begin VB.Label lblID 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      Height          =   315
      Left            =   4500
      TabIndex        =   4
      Top             =   210
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT REPAIR ORDER WITH PARTS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT ESTIMATE "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmCSMSQUESTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GetNewROno(XXX As Variant)
    Dim rsNewRO                                        As New ADODB.Recordset
    Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc")
    If Not rsNewRO.EOF And Not rsNewRO.BOF Then
        GetNewROno = Format(NumericVal(Mid$(rsNewRO!REP_OR, 3, 8)) + 1, "R-00000000")
    Else
        GetNewROno = "R-00000001"
    End If
    Set rsNewRO = Nothing
End Function

Sub PRINTING_TEST_ESTIMATE()
    If Null2String(lblDate.Caption) = "" Then
        gconDMIS.Execute "update CSMS_Esti_Hd set prin_dte = '" & Date & "' where id = " & lblID.Caption
        Call frmCSMSEstimateEntry.RECALL_STOREMEMVARS

        If CDbl(lblDISCOUNT.Caption) > 0 Then
            Call PRINTESTIDISC
        Else
            Call PRINTESTI
        End If
    Else
        If CDbl(lblDISCOUNT.Caption) > 0 Then
            Call PRINTESTIDISC
        Else
            Call PRINTESTI
        End If
    End If
End Sub

Sub PRINTING_TEST_REPAIR()
    txtRO.Text = GetNewROno(lblESTI.Caption)
    lblro.Caption = txtRO.Text
    picRO.Visible = True
    txtRO.SetFocus

End Sub

Sub PRINT_REPAIR()
    Screen.MousePointer = 11
    rptEstimate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEstimate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "estimatedisc_NEW.rpt", "{esti_hd.estimateno} = '" & lblESTI.Caption & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub PRINTESTIDISC()
    Screen.MousePointer = 11
    rptEstimate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEstimate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "estimatedisc.rpt", "{esti_hd.estimateno} = '" & lblESTI.Caption & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub PRINTESTI()
    Screen.MousePointer = 11
    rptEstimate.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptEstimate.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptEstimate, CSMS_REPORT_PATH & "estimate.rpt", "{esti_hd.estimateno} = '" & lblESTI.Caption & "'", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Sub UploadToRO()
    gconDMIS.Execute "update CSMS_Repor set" & _
                   " REP_OR = '" & txtRO.Text & "'," & _
                   " transtype = 'R'" & _
                   " where estimateno = '" & lblESTI.Caption & "'"

    gconDMIS.Execute "update CSMS_Ro_Det set" & _
                   " REP_OR = '" & txtRO.Text & "'," & _
                   " transtype = 'R'" & _
                   " where estimateno = '" & lblESTI.Caption & "'"

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
                   " RO_No = '" & txtRO.Text & "'," & _
                   " transtype = 'R'" & _
                   " where estimateno = '" & lblESTI.Caption & "'"

    gconDMIS.Execute "update CSMS_PMS_Job_Det set" & _
                   " REP_OR = '" & txtRO.Text & "'," & _
                   " transtype = 'R'" & _
                   " where estimateno = '" & lblESTI.Caption & "'"
End Sub

Sub CheckIfROExist(EXIST As Boolean)
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' And Rep_OR = '" & txtRO.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        EXIST = True
    Else
        EXIST = False
    End If
    Set rstmp = Nothing
End Sub

Private Sub cmdEXT_Click()
    picRO.Visible = False
End Sub

Private Sub cmdYES1_Click()
    QUESTION_TEST = "ESTIMATE"
    Call PRINTING_TEST_ESTIMATE
End Sub

Private Sub cmdYES2_Click()
    picRO.Visible = True
    txtRO.Text = ""
    txtRO.SetFocus
    Call PRINTING_TEST_REPAIR
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCSMSEstimateEntry.Enabled = True
End Sub

Private Sub txtRO_GotFocus()
    txtRO.BackColor = &HC0FFC0
End Sub

Private Sub txtRO_KeyPress(KeyAscii As Integer)
    Dim EXIST                                          As Boolean

    If KeyAscii = 13 Then
        If Not txtRO.Text = "" Then
            Call CheckIfROExist(EXIST)
            If Not EXIST Then
                If MsgBox("Are You Sure", vbQuestion + vbYesNo + vbDefaultButton1, "Upload Estimate To Repair Order") = vbYes Then
                    QUESTION_TEST = "REPAIR"

                    Call PRINT_REPAIR
                    Call UploadToRO

                    cmdEXT.Visible = False
                Else
                    txtRO.SetFocus
                End If
            Else
                MsgBox "RO Already Exist", vbCritical, "RO Already Exist"
                txtRO.SetFocus
            End If
        Else
            MsgBox "Enter A Repair RO", vbExclamation, "Incomplete Data"
            txtRO.SetFocus
        End If
    End If
End Sub

Private Sub txtRO_LostFocus()
    txtRO.BackColor = vbWhite
End Sub

