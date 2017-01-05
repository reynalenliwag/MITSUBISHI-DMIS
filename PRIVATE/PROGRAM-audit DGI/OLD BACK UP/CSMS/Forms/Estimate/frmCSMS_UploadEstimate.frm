VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_UploadEstimate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Estimate to RO"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_UploadEstimate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3930
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3060
      MouseIcon       =   "frmCSMS_UploadEstimate.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_UploadEstimate.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   1650
      Width           =   735
   End
   Begin VB.TextBox txtRO 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtESTNO 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   1
      Top             =   450
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpUploadDate 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   1230
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57409537
      CurrentDate     =   40035
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      Height          =   795
      Left            =   2340
      Picture         =   "frmCSMS_UploadEstimate.frx":1512
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label labCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order no"
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
      Index           =   60
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label labCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate no"
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
      Index           =   59
      Left            =   180
      TabIndex        =   5
      Top             =   510
      Width           =   975
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      _Version        =   655364
      _ExtentX        =   9128
      _ExtentY        =   609
      _StockProps     =   14
      Caption         =   "UPLOAD ESTIMATE TO REPAIR ORDER"
      ForeColor       =   14606302
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   0
   End
   Begin VB.Label labCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date to Upload"
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
      Index           =   61
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   1170
   End
End
Attribute VB_Name = "frmCSMS_UploadEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PutEstimateNo(ByVal xESTNO As String, FromForm As String)
Dim xFROMFROM                                   As String

Public Sub FillEstimateno(XXX As String, xxxFROMFROM As String)
    txtESTNO.Text = XXX
    xFROMFROM = xxxFROMFROM
End Sub

Private Sub cmdUpload_Click()
    Dim RSTMP                                          As New ADODB.Recordset
   
    Set RSTMP = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE REP_OR = '" & txtRO.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        MsgBox "Repair Order No Already Exist", vbExclamation, "CSMS"
        txtRO.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Upload this Estimate to Repair Order no: " & GenerateNewRONO & "", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    

    SQL_STATEMENT = "update CSMS_Repor set" & _
        " REP_OR = '" & txtRO & _
        "', transtype = 'R' " & _
        ", INSAMT = 0 " & _
        ", PARTLABOR = 0 " & _
        ", PARTPARTS = 0 " & _
        ", PARTMATERIALS = 0 " & _
        ", PARTACCESSORIES = 0 " & _
        " where estimateno = '" & txtESTNO & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UP", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtESTNO), "estimateno", "CSMS_REPOR"), "", "EST NO: " & txtESTNO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_Ro_Det set" & _
        " REP_OR = '" & txtRO & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtESTNO & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtESTNO), "estimateno", "CSMS_REPOR"), "", "EST NO: " & txtESTNO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
        " RO_No = '" & txtRO & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtESTNO & "'"

    gconDMIS.Execute "update CSMS_PMS_Job_Det set" & _
        " REP_OR = '" & txtRO & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtESTNO & "'"
    
    gconDMIS.Execute ("UPDATE CSMS_ESTHD SET " & _
        " DATE_UPLOAD = " & N2Str2Null(dtpUploadDate) & _
        ", UPLOAD_STATUS = 'Y' " & _
        ", RO_NO = " & N2Str2Null(txtRO) & _
        ", REP_OR = " & N2Str2Null(txtRO) & _
        " WHERE ESTIMATENO = " & N2Str2Null(txtESTNO) & "")
    
    gconDMIS.Execute ("DELETE FROM CSMS_RO_dET WHERE " & _
        " LIVIL <> 1 " & _
        " AND ESTIMATENO = " & N2Str2Null(txtESTNO) & "")
    
'    Dim ESTNO                 As String
'    Dim TOTJOBAMT             As Double
'    Dim TOTJOBDISC            As Double
'    Dim TOTJOBDISCVAL         As Double
'    Dim TOTJOBTAX             As Double
'    Dim JobComTotal           As Double
'    Dim JobSalesTotal         As Double
'    Dim JobWarTotal           As Double
'    Dim PartsComTotal           As Double
'    Dim PartsSalesTotal         As Double
'    Dim PartsWarTotal           As Double
'    Dim TOTPARTSAMT             As Double
'    Dim TOTPARTSDISC            As Double
'    Dim TOTPARTSDISCVAL         As Double
'    Dim TOTPARTSTAX             As Double
'    Dim ACCComTotal             As Double
'    Dim ACCSalesTotal           As Double
'    Dim ACCWarTotal             As Double
'
'    VTXTestimateno = xESTIMATENO
'    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
'    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
'
'    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where " & _
'        " EstimateNo = " & VTXTestimateno & _
'        " and livil = '1' " & _
'        " order by LINE_NO asc")
'    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
'        Screen.MousePointer = 11
'        rsRO_DET.MoveFirst
'        Do While Not rsRO_DET.EOF
'            If Null2String(rsRO_DET!wCode) = "C" Then
'                JobComTotal = K_JobComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
'            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
'            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
'            Else
'                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!DET_AMT)
'                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!Discount_2)
'                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
'                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!TAXVAL)
'            End If
'            rsRO_DET.MoveNext
'        Loop
'        Screen.MousePointer = 0
'    End If
'    Set rsRO_DET = Nothing
'    TOTJOBAMT = Round(TOTJOBAMT, 2)
'    TOTJOBDISC = Round(TOTJOBDISC, 2)
'    TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2)
'    TOTJOBTAX = Round(TOTJOBTAX, 2)
'
'    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
'    PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0
'    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
'    MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0
'    TOTACCAMT = 0: TOTACCDISC = 0: TOTACCDISCVAL = 0: TOTACCTAX = 0
'    ACCComTotal = 0: ACCSalesTotal = 0: ACCWarTotal = 0
'
'    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
'
'    gconDMIS.Execute "update CSMS_RepOr set" & _
'        " LABOR = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & _
'        ", l_amtvalue = " & Round(TOTJOBAMT, 2) & _
'        ", L_DISC = " & Round(TOTJOBDISCVAL, 2) & _
'        ", L_DISC2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & _
'        ", l_taxval = " & Round(TOTJOBTAX, 2) & _
'        ", l_discount = " & Round(TOTJOBDISC, 2) & _
'        ", PARTS = " & Round(TOTPARTSAMT - TOTPARTSTAX, 2) & _
'        ", P_AMTVALUE = " & Round(TOTPARTSAMT, 2) & ", P_DISC = " & Round(TOTPARTSDISCVAL, 2) & _
'        ", P_DISC2 = " & Round(TOTPARTSDISC * (VAT_RATE / 100), 2) & ", P_TAXVAL = " & Round(TOTPARTSTAX) & _
'        ", MATERIAL = " & Round(TOTMATAMT - TOTMATTAX, 2) & _
'        ", M_AMTVALUE = " & Round(TOTMATAMT, 2) & ", M_DISC = " & Round(TOTMATDISCVAL, 2) & _
'        ", M_DISC2 = " & Round(TOTMATDISC * (VAT_RATE / 100), 2) & ", M_TAXVAL = " & Round(TOTMATTAX) & _
'        ", ACCESSORIES = " & Round(TOTACCAMT - TOTACCTAX, 2) & _
'        ", A_AMTVALUE = " & Round(TOTACCAMT, 2) & ", A_DISC = " & Round(TOTACCDISCVAL, 2) & _
'        ", A_DISC2 = " & Round(TOTACCDISC * (VAT_RATE / 100), 2) & ", A_TAXVAL = " & Round(TOTACCTAX) & _
'        ", AMOUNT = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        ", ROVAT = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & _
'        ", WL_AMOUNT = " & 0 & _
'        ", RO_AMOUNT = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
'        " where " & FilterType
        
    Call ShowSuccessFullyUpdated
    
    RaiseEvent PutEstimateNo(txtESTNO, xFROMFROM)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    txtRO.Text = GenerateNewRONO
    dtpUploadDate.Value = Date
End Sub

Function GenerateNewRONO() As String
    Dim RSTMP                                           As New ADODB.Recordset
    Dim vRO                                             As String
    Set RSTMP = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR ORDER BY REP_OR DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        vRO = Mid(RSTMP!REP_OR, 3, 8) + 1
        GenerateNewRONO = "R-" & Format(vRO, "00000000")
    Else
        GenerateNewRONO = "R-00000001'"
    End If
    Set RSTMP = Nothing
End Function
