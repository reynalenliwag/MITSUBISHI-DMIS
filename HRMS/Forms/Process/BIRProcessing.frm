VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmHRMSBIRProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Processing"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9510
   ControlBox      =   0   'False
   ForeColor       =   &H00D8E9EC&
   Icon            =   "BIRProcessing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   9510
   Begin VB.ComboBox cboSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   225
      Width           =   8415
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   885
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
      Left            =   8550
      MouseIcon       =   "BIRProcessing.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "BIRProcessing.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel Processing"
      Top             =   1245
      Width           =   885
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Process"
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
      Left            =   7680
      MouseIcon       =   "BIRProcessing.frx":08FE
      MousePointer    =   99  'Custom
      Picture         =   "BIRProcessing.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process"
      Top             =   1245
      Width           =   885
   End
   Begin Crystal.CrystalReport rptPrintYTD 
      Left            =   600
      Top             =   1665
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComDlg.CommonDialog cmdDialogPIS 
      Left            =   60
      Top             =   1635
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   9405
      TabIndex        =   4
      Top             =   585
      Width           =   9405
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   5
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
            TabIndex        =   6
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   7
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   8
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
            MICON           =   "BIRProcessing.frx":0DBE
         End
      End
      Begin wizProgBar.Prg progYTDProcessing 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   300
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   556
         Picture         =   "BIRProcessing.frx":0DDA
         BackColor       =   14215660
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "BIRProcessing.frx":0DF6
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
      Begin VB.Label labEmpNo 
         BackColor       =   &H00D8E9EC&
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
         TabIndex        =   9
         Top             =   30
         Width           =   3585
      End
   End
End
Attribute VB_Name = "frmHRMSBIRProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsYTDDETAILS, rsHeader                                 As ADODB.Recordset
Attribute rsYTDDETAILS.VB_VarUserMemId = 1073938432
Attribute rsHeader.VB_VarUserMemId = 1073938432
Dim rsDetails71, rsControls71, rsDetails73                            As ADODB.Recordset
Attribute rsDetails71.VB_VarUserMemId = 1073938435
Attribute rsControls71.VB_VarUserMemId = 1073938435
Attribute rsDetails73.VB_VarUserMemId = 1073938435
Dim rsControls73, rsDetails74, rsControls74                           As ADODB.Recordset
Attribute rsControls73.VB_VarUserMemId = 1073938438
Attribute rsDetails74.VB_VarUserMemId = 1073938438
Attribute rsControls74.VB_VarUserMemId = 1073938438
Dim ToDate, THEBIRCS, EMPLOYER_TIN, FILNAME                           As String
Attribute ToDate.VB_VarUserMemId = 1073938441
Attribute THEBIRCS.VB_VarUserMemId = 1073938441
Attribute EMPLOYER_TIN.VB_VarUserMemId = 1073938441
Attribute FILNAME.VB_VarUserMemId = 1073938441
Dim dagos                                                             As Boolean
Attribute dagos.VB_VarUserMemId = 1073938445

Sub TERMINATED_PROC()
    Dim YY                                                            As String
    YY = cboyear.Text
    Dim vSchedule_num, vFtype_code, vTin_Empyr, vBranch_Code_Emplyr   As String
    Dim vRetrn_Period, vTin, vBranch_Code, vLast_Name                 As String
    Dim vSeq_num                                                      As Integer
    Dim vFirst_Name, vMiddle_Name, vEmployment_From, vEmployment_To   As String

    Dim vPres_Nontax_13th_month, vPres_Nontax_SSS_Etc, vPres_Nontax_Salaries As Double
    Dim vPres_Taxable_13th_month, vPres_Taxable_Salaries, vExmpn_Amt  As Double
    Dim vPremium_Paid, vTax_Due, vPres_Tax_wthld, vAmt_Wthld_Dec      As Double
    Dim vOver_Wthld, vActual_Amt_Wthld                                As Double

    Dim vTOTALPres_Nontax_13th_month, vTOTALPres_Nontax_SSS_Etc       As Double
    Dim vTOTALPres_Nontax_Salaries, vTOTALPres_Taxable_13th_month     As Double
    Dim vTOTALPres_Taxable_Salaries, vTOTALExmpn_Amt                  As Double
    Dim vTOTALPremium_Paid, vTOTALTax_Due                             As Double
    Dim vTOTALPres_Tax_wthld, vTOTALAmt_Wthld_Dec                     As Double
    Dim vTOTALOver_Wthld, vTOTALActual_Amt_Wthld                      As Double

    Dim vCSVSchedule_num, vCSVFtype_code, vCSVTin_Empyr, vCSVBranch_Code_Emplyr As String
    Dim vCSVRetrn_Period                                              As String
    Dim vCSVSeq_num                                                   As Integer
    Dim vCSVTin, vCSVBranch_Code, vCSVLast_Name, vCSVFirst_Name, vCSVMiddle_Name As String
    Dim vCSVEmployment_From, vCSVEmployment_To                        As String
    Dim vCSVPres_Nontax_13th_month, vCSVPres_Nontax_SSS_Etc           As Double
    Dim vCSVPres_Nontax_Salaries, vCSVPres_Taxable_13th_month         As Double
    Dim vCSVPres_Taxable_Salaries, vCSVExmpn_Amt                      As Double
    Dim vCSVPremium_Paid, vCSVTax_Due                                 As Double
    Dim vCSVPres_Tax_wthld, vCSVAmt_Wthld_Dec                         As Double
    Dim vCSVOver_Wthld, vCSVActual_Amt_Wthld                          As Double

    Dim I, CNT                                                        As Integer
    Dim schedFName                                                    As String
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        I = 1: CNT = 0
        schedFName = EMPLOYER_TIN & ".s71"
        vSchedule_num = "'D7.1'"
        vFtype_code = "'1604CF'"
        vTin_Empyr = "'" & EMPLOYER_TIN & "'"
        vBranch_Code_Emplyr = "NULL"
        vRetrn_Period = "'01/01/" & YY & "'"

        vCSVSchedule_num = "D7.1"
        vCSVFtype_code = "1604CF"
        vCSVTin_Empyr = "005532084"
        vCSVBranch_Code_Emplyr = ""
        vCSVRetrn_Period = "01/01/" & YY
        On Error Resume Next
        Dim MYPATH, PAYLNAME                                          As String
        MYPATH = App.Path
        cmdDialogPIS.FILTER = "Schedule Files (*.s71)|*.s71"
        cmdDialogPIS.FilterIndex = 1
        cmdDialogPIS.DefaultExt = "S71"
        PAYLNAME = cmdDialogPIS.Filename
        If MYPATH <> "\" Then
            cmdDialogPIS.Filename = MYPATH & "\" & cmdDialogPIS.Filename
        End If
        If PAYLNAME = "" Then
            cmdDialogPIS.Filename = schedFName
        End If
        cmdDialogPIS.Action = 2
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        FILNAME = cmdDialogPIS.Filename
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        If dagos = True Then
            Open schedFName For Output As #1
            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
            Print #1, "1604CF" & "," & EMPLOYER_TIN & ",," & "01/01/" & YY
            Do While Not rsEmpInfo.EOF
                If Null2String(rsEmpInfo!RESIGNED) <> "" Then
                    If YEAR(Null2String(rsEmpInfo!RESIGNED)) = YY Then
                        Set rsYTDDETAILS = New ADODB.Recordset
                        rsYTDDETAILS.Open "select * from HRMS_YTDDetails where YEER = '" & YY & "' AND empno ='" & rsEmpInfo!EMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                            vSeq_num = I
                            vTin = N2Str2Null(Null2String(rsEmpInfo!tinno))
                            vBranch_Code = N2Str2Null(Null2String(rsEmpInfo!EMPNO))
                            vLast_Name = N2Str2Null(Null2String(rsEmpInfo!lastname))
                            vFirst_Name = N2Str2Null(Null2String(rsEmpInfo!FIRSTNAME))
                            vMiddle_Name = N2Str2Null(Left(Null2String(rsEmpInfo!MIDDLENAME), 1))
                            vEmployment_From = N2Str2Null(Format(Null2String(rsEmpInfo!DateHired), "MM/DD/YYYY"))
                            vEmployment_To = N2Str2Null(Format(Null2String(rsEmpInfo!RESIGNED), "MM/DD/YYYY"))
                            vPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)

                            vPres_Taxable_13th_month = 0
                            vPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vPremium_Paid = 0
                            vTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            If vTax_Due > vPres_Tax_wthld Then
                                vAmt_Wthld_Dec = Format(vTax_Due - vPres_Tax_wthld, MAXIMUM_DIGIT)
                            Else
                                vAmt_Wthld_Dec = 0
                            End If
                            If vTax_Due < vPres_Tax_wthld Then
                                vOver_Wthld = Format(vPres_Tax_wthld - vTax_Due, MAXIMUM_DIGIT)
                            Else
                                vOver_Wthld = 0
                            End If
                            If vAmt_Wthld_Dec > vOver_Wthld Then
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld + vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            Else
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld - vOver_Wthld, MAXIMUM_DIGIT)
                            End If

                            'INITIALIZE TOTAL VALUES
                            vTOTALPres_Nontax_13th_month = vTOTALPres_Nontax_13th_month + vPres_Nontax_13th_month
                            vTOTALPres_Nontax_SSS_Etc = vTOTALPres_Nontax_SSS_Etc + vPres_Nontax_SSS_Etc
                            vTOTALPres_Nontax_Salaries = vTOTALPres_Nontax_Salaries + vPres_Nontax_Salaries
                            vTOTALPres_Taxable_13th_month = vTOTALPres_Taxable_13th_month + vPres_Taxable_13th_month
                            vTOTALPres_Taxable_Salaries = vTOTALPres_Taxable_Salaries + vPres_Taxable_Salaries
                            vTOTALExmpn_Amt = vTOTALExmpn_Amt + vExmpn_Amt
                            vTOTALPremium_Paid = vTOTALPremium_Paid + vPremium_Paid
                            vTOTALTax_Due = vTOTALTax_Due + vTax_Due
                            vTOTALPres_Tax_wthld = vTOTALPres_Tax_wthld + vPres_Tax_wthld
                            vTOTALAmt_Wthld_Dec = vTOTALAmt_Wthld_Dec + vAmt_Wthld_Dec
                            vTOTALOver_Wthld = vTOTALOver_Wthld + vOver_Wthld
                            vTOTALActual_Amt_Wthld = vTOTALActual_Amt_Wthld + vActual_Amt_Wthld

                            'INITIALIZE COMMA SEPARATED VALUE
                            vCSVSeq_num = I
                            vCSVTin = Null2String(rsEmpInfo!tinno)
                            vCSVBranch_Code = Null2String(rsEmpInfo!EMPNO)
                            vCSVLast_Name = Null2String(rsEmpInfo!lastname)
                            vCSVFirst_Name = Null2String(rsEmpInfo!FIRSTNAME)
                            vCSVMiddle_Name = Left(Null2String(rsEmpInfo!MIDDLENAME), 1)
                            vCSVEmployment_From = Format(Null2String(rsEmpInfo!DateHired), "MM/DD/YYYY")
                            vCSVEmployment_To = Format(Null2String(rsEmpInfo!RESIGNED), "MM/DD/YYYY")
                            vCSVPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)
                            vCSVPres_Taxable_13th_month = 0
                            vCSVPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vCSVExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vCSVPremium_Paid = 0
                            vCSVTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vCSVPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            vCSVAmt_Wthld_Dec = Format(vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            vCSVOver_Wthld = Format(vOver_Wthld, MAXIMUM_DIGIT)
                            vCSVActual_Amt_Wthld = Format(vActual_Amt_Wthld, MAXIMUM_DIGIT)

                            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vCSVSeq_num & "," & vCSVTin & "," & vCSVBranch_Code & "," & vCSVLast_Name & "," & vCSVFirst_Name & "," & vCSVMiddle_Name & "," & vCSVEmployment_From & "," & vCSVEmployment_To & "," & vCSVPres_Nontax_13th_month & "," & vCSVPres_Nontax_SSS_Etc & "," & vCSVPres_Nontax_Salaries & "," & vCSVPres_Taxable_13th_month & "," & vCSVPres_Taxable_Salaries & "," & vCSVExmpn_Amt & "," & vCSVPremium_Paid & "," & vCSVTax_Due & "," & vCSVPres_Tax_wthld & "," & vCSVAmt_Wthld_Dec & "," & vCSVOver_Wthld & "," & vCSVActual_Amt_Wthld
                            Set rsDetails71 = New ADODB.Recordset
                            rsDetails71.Open "select * from HRMS_Details71 where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'", gconDMIS
                            If Not rsDetails71.EOF And Not rsDetails71.BOF Then
                                gconDMIS.Execute "update HRMS_Details71 set " & _
                                                 "schedule_num = " & vSchedule_num & ", " & _
                                                 "ftype_code = " & vFtype_code & ", " & _
                                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                                 "retrn_period = " & vRetrn_Period & ", " & _
                                                 "seq_num = " & vSeq_num & ", " & _
                                                 "tin =" & vTin & ", " & _
                                                 "last_name = " & vLast_Name & ", " & _
                                                 "first_name = " & vFirst_Name & ", " & _
                                                 "middle_name = " & vMiddle_Name & ", " & _
                                                 "employment_from = " & vEmployment_From & ", " & _
                                                 "employment_to = " & vEmployment_To & ", " & _
                                                 "pres_nontax_13th_month = " & vPres_Nontax_13th_month & ", " & _
                                                 "pres_nontax_sss_etc = " & vPres_Nontax_SSS_Etc & ", " & _
                                                 "pres_nontax_salaries = " & vPres_Nontax_Salaries & ", " & _
                                                 "pres_taxable_13th_month = " & vPres_Taxable_13th_month & ", " & _
                                                 "pres_taxable_salaries = " & vPres_Taxable_Salaries & ", " & _
                                                 "exmpn_amt = " & vExmpn_Amt & ", " & _
                                                 "premium_paid = " & vPremium_Paid & ", " & _
                                                 "tax_due = " & vTax_Due & ", " & _
                                                 "pres_tax_wthld = " & vPres_Tax_wthld & ", " & _
                                                 "amt_wthld_dec = " & vAmt_Wthld_Dec & ", over_wthld = " & vOver_Wthld & ", actual_amt_wthld = " & vActual_Amt_Wthld & _
                                               " where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'"
                            Else
                                gconDMIS.Execute "insert into HRMS_Details71 " & _
                                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,seq_num,tin,branch_code,last_name,first_name,middle_name,employment_from,employment_to,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,exmpn_amt,premium_paid,tax_due,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ", " & vSeq_num & ", " & vTin & ", " & vBranch_Code & ", " & vLast_Name & _
                                                 ", " & vFirst_Name & ", " & vMiddle_Name & _
                                                 ", " & vEmployment_From & ", " & vEmployment_To & _
                                                 ", " & vPres_Nontax_13th_month & ", " & vPres_Nontax_SSS_Etc & _
                                                 ", " & vPres_Nontax_Salaries & ", " & vPres_Taxable_13th_month & _
                                                 ", " & vPres_Taxable_Salaries & ", " & vExmpn_Amt & ", " & vPremium_Paid & ", " & vTax_Due & ", " & vPres_Tax_wthld & ", " & vAmt_Wthld_Dec & ", " & vOver_Wthld & ", " & vActual_Amt_Wthld & ")"
                            End If
                            'Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vCSVSeq_num & "," & vCSVTin & "," & vCSVBranch_Code & "," & vCSVLast_Name & "," & vCSVFirst_Name & "," & vCSVMiddle_Name & "," & vCSVEmployment_From & "," & vCSVEmployment_To & "," & vCSVPres_Nontax_13th_month & "," & vCSVPres_Nontax_SSS_Etc & "," & vCSVPres_Nontax_Salaries & "," & vCSVPres_Taxable_13th_month & "," & vCSVPres_Taxable_Salaries & "," & vCSVExmpn_Amt & "," & vCSVPremium_Paid & "," & vCSVTax_Due & "," & vCSVPres_Tax_wthld & "," & vCSVAmt_Wthld_Dec & "," & vCSVOver_Wthld & "," & vCSVActual_Amt_Wthld
                            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
                            I = I + 1
                        End If
                    End If
                End If
                CNT = CNT + 1
                progYTDProcessing.Value = (CNT / rsEmpInfo.RecordCount) * 100
                labEmpNo.Caption = Int(progYTDProcessing.Value) & "%"
                DoEvents
                rsEmpInfo.MoveNext
            Loop
            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vTOTALPres_Nontax_13th_month & "," & vTOTALPres_Nontax_SSS_Etc & "," & vTOTALPres_Nontax_Salaries & "," & vTOTALPres_Taxable_13th_month & "," & vTOTALPres_Taxable_Salaries & "," & vTOTALExmpn_Amt & "," & vTOTALPremium_Paid & "," & vTOTALTax_Due & "," & vTOTALPres_Tax_wthld & "," & vTOTALAmt_Wthld_Dec & "," & vTOTALOver_Wthld & "," & vTOTALActual_Amt_Wthld
            Set rsControls71 = New ADODB.Recordset
            rsControls71.Open "select * from HRMS_Controls71 where ftype_code = " & vFtype_code, gconDMIS
            If Not rsControls71.EOF And Not rsControls71.BOF Then
                gconDMIS.Execute "update HRMS_Controls71 set " & _
                                 "schedule_num = " & vSchedule_num & ", " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "pres_nontax_13th_month = " & vTOTALPres_Nontax_13th_month & ", " & _
                                 "pres_nontax_sss_etc = " & vTOTALPres_Nontax_SSS_Etc & ", " & _
                                 "pres_nontax_salaries = " & vTOTALPres_Nontax_Salaries & ", " & _
                                 "pres_taxable_13th_month = " & vTOTALPres_Taxable_13th_month & ", " & _
                                 "pres_taxable_salaries = " & vTOTALPres_Taxable_Salaries & ", " & _
                                 "exmpn_amt = " & vTOTALExmpn_Amt & ", " & _
                                 "premium_paid = " & vTOTALPremium_Paid & ", " & _
                                 "tax_due = " & vTOTALTax_Due & ", " & _
                                 "pres_tax_wthld = " & vTOTALPres_Tax_wthld & ", " & _
                                 "amt_wthld_dec = " & vTOTALAmt_Wthld_Dec & ", over_wthld = " & vTOTALOver_Wthld & ", actual_amt_wthld = " & vTOTALActual_Amt_Wthld & _
                               " where ftype_code = " & vFtype_code
            Else
                gconDMIS.Execute "insert into HRMS_Controls71 " & _
                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,exmpn_amt,premium_paid,tax_due,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & _
                                 ", " & vTOTALPres_Nontax_13th_month & ", " & vTOTALPres_Nontax_SSS_Etc & _
                                 ", " & vTOTALPres_Nontax_Salaries & ", " & vTOTALPres_Taxable_13th_month & _
                                 ", " & vTOTALPres_Taxable_Salaries & ", " & vTOTALExmpn_Amt & ", " & vTOTALPremium_Paid & ", " & vTOTALTax_Due & ", " & vTOTALPres_Tax_wthld & ", " & vTOTALAmt_Wthld_Dec & ", " & vTOTALOver_Wthld & ", " & vTOTALActual_Amt_Wthld & ")"
            End If
            Set rsHeader = New ADODB.Recordset
            rsHeader.Open "select * from HRMS_Header where ftype_code = " & vFtype_code, gconDMIS
            If Not rsHeader.EOF And Not rsHeader.BOF Then
                gconDMIS.Execute "update HRMS_Header set " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin = " & vTin_Empyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "branch_code = " & vBranch_Code_Emplyr & _
                               " where ftype_code = " & vFtype_code
            Else
                gconDMIS.Execute "insert into HRMS_Header " & _
                                 "(ftype_code,tin,branch_code,retrn_period)" & _
                               " values (" & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ")"
            End If
            Close #1
        Else
            MsgSpeechBox "Process was Cancelled!"
        End If
    Else
        ShowNoRecord
    End If
End Sub

Sub WITH_NO_PREV()
    Dim YY                                                            As String
    YY = cboyear.Text
    Dim vSchedule_num, vFtype_code, vTin_Empyr, vBranch_Code_Emplyr, vRetrn_Period As String
    Dim vSeq_num                                                      As Integer
    Dim vTin, vBranch_Code, vLast_Name, vFirst_Name, vMiddle_Name     As String
    Dim vPres_Nontax_13th_month, vPres_Nontax_SSS_Etc, vPres_Nontax_Salaries As Double
    Dim vPres_Taxable_13th_month, vPres_Taxable_Salaries              As Double
    Dim vExmpn_Amt, vPremium_Paid, vTax_Due, vPres_Tax_wthld          As Double
    Dim vAmt_Wthld_Dec, vOver_Wthld, vActual_Amt_Wthld                As Double

    Dim vTOTALPres_Nontax_13th_month, vTOTALPres_Nontax_SSS_Etc       As Double
    Dim vTOTALPres_Nontax_Salaries, vTOTALPres_Taxable_13th_month     As Double
    Dim vTOTALPres_Taxable_Salaries, vTOTALExmpn_Amt                  As Double
    Dim vTOTALPremium_Paid, vTOTALTax_Due                             As Double
    Dim vTOTALPres_Tax_wthld, vTOTALAmt_Wthld_Dec                     As Double
    Dim vTOTALOver_Wthld, vTOTALActual_Amt_Wthld                      As Double

    Dim vCSVSchedule_num, vCSVFtype_code, vCSVTin_Empyr               As String
    Dim vCSVBranch_Code_Emplyr, vCSVRetrn_Period                      As String
    Dim vCSVSeq_num                                                   As Integer
    Dim vCSVTin, vCSVBranch_Code, vCSVLast_Name                       As String
    Dim vCSVFirst_Name, vCSVMiddle_Name                               As String
    Dim vCSVPres_Nontax_13th_month, vCSVPres_Nontax_SSS_Etc           As Double
    Dim vCSVPres_Nontax_Salaries, vCSVPres_Taxable_13th_month         As Double
    Dim vCSVPres_Taxable_Salaries, vCSVExmpn_Amt                      As Double
    Dim vCSVPremium_Paid, vCSVTax_Due                                 As Double
    Dim vCSVPres_Tax_wthld, vCSVAmt_Wthld_Dec                         As Double
    Dim vCSVOver_Wthld, vCSVActual_Amt_Wthld                          As Double

    Dim I, CNT                                                        As Integer
    Dim schedFName                                                    As String
    Dim Sige                                                          As Boolean
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        I = 1: CNT = 0
        schedFName = EMPLOYER_TIN & ".s73"
        vSchedule_num = "'D7.3'"
        vFtype_code = "'1604CF'"
        vTin_Empyr = "'" & EMPLOYER_TIN & "'"
        vBranch_Code_Emplyr = "NULL"
        vRetrn_Period = "'01/01/" & YY & "'"

        vCSVSchedule_num = "D7.3"
        vCSVFtype_code = "1604CF"
        vCSVTin_Empyr = EMPLOYER_TIN
        vCSVBranch_Code_Emplyr = ""
        vCSVRetrn_Period = "01/01/" & YY
        On Error Resume Next
        Dim MYPATH                                                    As String
        Dim PAYLNAME                                                  As String
        MYPATH = App.Path
        cmdDialogPIS.FILTER = "Schedule Files (*.s73)|*.s73"
        cmdDialogPIS.FilterIndex = 1
        cmdDialogPIS.DefaultExt = "S73"
        PAYLNAME = cmdDialogPIS.Filename
        If MYPATH <> "\" Then
            cmdDialogPIS.Filename = MYPATH & "\" & cmdDialogPIS.Filename
        End If
        If PAYLNAME = "" Then
            cmdDialogPIS.Filename = schedFName
        End If

        cmdDialogPIS.Action = 2
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        FILNAME = cmdDialogPIS.Filename
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        If dagos = True Then
            Open schedFName For Output As #1
            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
            Print #1, "1604CF" & "," & EMPLOYER_TIN & ",," & "01/01/" & YY
            Do While Not rsEmpInfo.EOF
                If Null2String(rsEmpInfo!RESIGNED) = "" Then
                    Sige = True
                Else
                    Sige = False
                    If YEAR(Null2String(rsEmpInfo!RESIGNED)) <> YY Then
                        Sige = True
                    End If
                End If
                If Sige = True Then
                    If Null2String(rsEmpInfo!withprevious) = "N" Then
                        Set rsYTDDETAILS = New ADODB.Recordset
                        rsYTDDETAILS.Open "select * from HRMS_YTDDetails where YEER = '" & YY & "' AND empno ='" & rsEmpInfo!EMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                            vSeq_num = I
                            vTin = N2Str2Null(Null2String(rsEmpInfo!tinno))
                            vBranch_Code = N2Str2Null(Null2String(rsEmpInfo!EMPNO))
                            vLast_Name = N2Str2Null(Null2String(rsEmpInfo!lastname))
                            vFirst_Name = N2Str2Null(Null2String(rsEmpInfo!FIRSTNAME))
                            vMiddle_Name = N2Str2Null(Left(Null2String(rsEmpInfo!MIDDLENAME), 1))
                            vPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)

                            vPres_Taxable_13th_month = 0
                            vPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vPremium_Paid = 0
                            vTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            If vTax_Due > vPres_Tax_wthld Then
                                vAmt_Wthld_Dec = Format(vTax_Due - vPres_Tax_wthld, MAXIMUM_DIGIT)
                            Else
                                vAmt_Wthld_Dec = 0
                            End If
                            If vTax_Due < vPres_Tax_wthld Then
                                vOver_Wthld = Format(vPres_Tax_wthld - vTax_Due, MAXIMUM_DIGIT)
                            Else
                                vOver_Wthld = 0
                            End If
                            If vAmt_Wthld_Dec > vOver_Wthld Then
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld + vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            Else
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld - vOver_Wthld, MAXIMUM_DIGIT)
                            End If

                            'INITIALIZE TOTAL VALUES
                            vTOTALPres_Nontax_13th_month = vTOTALPres_Nontax_13th_month + vPres_Nontax_13th_month
                            vTOTALPres_Nontax_SSS_Etc = vTOTALPres_Nontax_SSS_Etc + vPres_Nontax_SSS_Etc
                            vTOTALPres_Nontax_Salaries = vTOTALPres_Nontax_Salaries + vPres_Nontax_Salaries
                            vTOTALPres_Taxable_13th_month = vTOTALPres_Taxable_13th_month + vPres_Taxable_13th_month
                            vTOTALPres_Taxable_Salaries = vTOTALPres_Taxable_Salaries + vPres_Taxable_Salaries
                            vTOTALExmpn_Amt = vTOTALExmpn_Amt + vExmpn_Amt
                            vTOTALPremium_Paid = vTOTALPremium_Paid + vPremium_Paid
                            vTOTALTax_Due = vTOTALTax_Due + vTax_Due
                            vTOTALPres_Tax_wthld = vTOTALPres_Tax_wthld + vPres_Tax_wthld
                            vTOTALAmt_Wthld_Dec = vTOTALAmt_Wthld_Dec + vAmt_Wthld_Dec
                            vTOTALOver_Wthld = vTOTALOver_Wthld + vOver_Wthld
                            vTOTALActual_Amt_Wthld = vTOTALActual_Amt_Wthld + vActual_Amt_Wthld

                            'INITIALIZE COMMA SEPARATED VALUE
                            vCSVSeq_num = I
                            vCSVTin = Null2String(rsEmpInfo!tinno)
                            vCSVBranch_Code = Null2String(rsEmpInfo!EMPNO)
                            vCSVLast_Name = Null2String(rsEmpInfo!lastname)
                            vCSVFirst_Name = Null2String(rsEmpInfo!FIRSTNAME)
                            vCSVMiddle_Name = Left(Null2String(rsEmpInfo!MIDDLENAME), 1)
                            vCSVPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)
                            vCSVPres_Taxable_13th_month = 0
                            vCSVPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vCSVExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vCSVPremium_Paid = 0
                            vCSVTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vCSVPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            vCSVAmt_Wthld_Dec = Format(vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            vCSVOver_Wthld = Format(vOver_Wthld, MAXIMUM_DIGIT)
                            vCSVActual_Amt_Wthld = Format(vActual_Amt_Wthld, MAXIMUM_DIGIT)

                            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vCSVSeq_num & "," & vCSVTin & "," & vCSVBranch_Code & "," & vCSVLast_Name & "," & vCSVFirst_Name & "," & vCSVMiddle_Name & "," & vCSVPres_Nontax_13th_month & "," & vCSVPres_Nontax_SSS_Etc & "," & vCSVPres_Nontax_Salaries & "," & vCSVPres_Taxable_13th_month & "," & vCSVPres_Taxable_Salaries & "," & vCSVExmpn_Amt & "," & vCSVPremium_Paid & "," & vCSVTax_Due & "," & vCSVPres_Tax_wthld & "," & vCSVAmt_Wthld_Dec & "," & vCSVOver_Wthld & "," & vCSVActual_Amt_Wthld

                            Set rsDetails73 = New ADODB.Recordset
                            rsDetails73.Open "select * from HRMS_Details73 where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'", gconDMIS
                            If Not rsDetails73.EOF And Not rsDetails73.BOF Then
                                gconDMIS.Execute "update HRMS_Details73 set " & _
                                                 "schedule_num = " & vSchedule_num & ", " & _
                                                 "ftype_code = " & vFtype_code & ", " & _
                                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                                 "retrn_period = " & vRetrn_Period & ", " & _
                                                 "seq_num = " & vSeq_num & ", " & _
                                                 "tin =" & vTin & ", " & _
                                                 "last_name = " & vLast_Name & ", " & _
                                                 "first_name = " & vFirst_Name & ", " & _
                                                 "middle_name = " & vMiddle_Name & ", " & _
                                                 "pres_nontax_13th_month = " & vPres_Nontax_13th_month & ", " & _
                                                 "pres_nontax_sss_etc = " & vPres_Nontax_SSS_Etc & ", " & _
                                                 "pres_nontax_salaries = " & vPres_Nontax_Salaries & ", " & _
                                                 "pres_taxable_13th_month = " & vPres_Taxable_13th_month & ", " & _
                                                 "pres_taxable_salaries = " & vPres_Taxable_Salaries & ", " & _
                                                 "exmpn_amt = " & vExmpn_Amt & ", " & _
                                                 "premium_paid = " & vPremium_Paid & ", " & _
                                                 "tax_due = " & vTax_Due & ", " & _
                                                 "pres_tax_wthld = " & vPres_Tax_wthld & ", " & _
                                                 "amt_wthld_dec = " & vAmt_Wthld_Dec & ", over_wthld = " & vOver_Wthld & ", actual_amt_wthld = " & vActual_Amt_Wthld & _
                                               " where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'"
                            Else
                                gconDMIS.Execute "insert into HRMS_Details73 " & _
                                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,seq_num,tin,branch_code,last_name,first_name,middle_name,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,exmpn_amt,premium_paid,tax_due,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ", " & vSeq_num & ", " & vTin & ", " & vBranch_Code & ", " & vLast_Name & _
                                                 ", " & vFirst_Name & ", " & vMiddle_Name & _
                                                 ", " & vPres_Nontax_13th_month & ", " & vPres_Nontax_SSS_Etc & _
                                                 ", " & vPres_Nontax_Salaries & ", " & vPres_Taxable_13th_month & _
                                                 ", " & vPres_Taxable_Salaries & ", " & vExmpn_Amt & ", " & vPremium_Paid & ", " & vTax_Due & ", " & vPres_Tax_wthld & ", " & vAmt_Wthld_Dec & ", " & vOver_Wthld & ", " & vActual_Amt_Wthld & ")"
                            End If
                            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
                            I = I + 1
                        End If
                    End If
                End If
                CNT = CNT + 1
                progYTDProcessing.Value = (CNT / rsEmpInfo.RecordCount) * 100
                labEmpNo.Caption = Int(progYTDProcessing.Value) & "%"
                DoEvents
                rsEmpInfo.MoveNext
            Loop
            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vTOTALPres_Nontax_13th_month & "," & vTOTALPres_Nontax_SSS_Etc & "," & vTOTALPres_Nontax_Salaries & "," & vTOTALPres_Taxable_13th_month & "," & vTOTALPres_Taxable_Salaries & "," & vTOTALExmpn_Amt & "," & vTOTALPremium_Paid & "," & vTOTALTax_Due & "," & vTOTALPres_Tax_wthld & "," & vTOTALAmt_Wthld_Dec & "," & vTOTALOver_Wthld & "," & vTOTALActual_Amt_Wthld
            Set rsControls73 = New ADODB.Recordset
            rsControls73.Open "select * from HRMS_Controls73 where schedule_num = " & vSchedule_num, gconDMIS
            If Not rsControls73.EOF And Not rsControls73.BOF Then
                gconDMIS.Execute "update HRMS_Controls73 set " & _
                                 "schedule_num = " & vSchedule_num & ", " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "pres_nontax_13th_month = " & vTOTALPres_Nontax_13th_month & ", " & _
                                 "pres_nontax_sss_etc = " & vTOTALPres_Nontax_SSS_Etc & ", " & _
                                 "pres_nontax_salaries = " & vTOTALPres_Nontax_Salaries & ", " & _
                                 "pres_taxable_13th_month = " & vTOTALPres_Taxable_13th_month & ", " & _
                                 "pres_taxable_salaries = " & vTOTALPres_Taxable_Salaries & ", " & _
                                 "exmpn_amt = " & vTOTALExmpn_Amt & ", " & _
                                 "premium_paid = " & vTOTALPremium_Paid & ", " & _
                                 "tax_due = " & vTOTALTax_Due & ", " & _
                                 "pres_tax_wthld = " & vTOTALPres_Tax_wthld & ", " & _
                                 "amt_wthld_dec = " & vTOTALAmt_Wthld_Dec & ", over_wthld = " & vTOTALOver_Wthld & ", actual_amt_wthld = " & vTOTALActual_Amt_Wthld & _
                               " where schedule_num = " & vSchedule_num
            Else
                gconDMIS.Execute "insert into HRMS_Controls73 " & _
                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,exmpn_amt,premium_paid,tax_due,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & _
                                 ", " & vTOTALPres_Nontax_13th_month & ", " & vTOTALPres_Nontax_SSS_Etc & _
                                 ", " & vTOTALPres_Nontax_Salaries & ", " & vTOTALPres_Taxable_13th_month & _
                                 ", " & vTOTALPres_Taxable_Salaries & ", " & vTOTALExmpn_Amt & ", " & vTOTALPremium_Paid & ", " & vTOTALTax_Due & ", " & vTOTALPres_Tax_wthld & ", " & vTOTALAmt_Wthld_Dec & ", " & vTOTALOver_Wthld & ", " & vTOTALActual_Amt_Wthld & ")"
            End If
            Set rsHeader = New ADODB.Recordset
            rsHeader.Open "select * from HRMS_Header where ftype_code = " & vFtype_code, gconDMIS
            If Not rsHeader.EOF And Not rsHeader.BOF Then
                gconDMIS.Execute "update HRMS_Header set " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin = " & vTin_Empyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "branch_code = " & vBranch_Code_Emplyr & _
                               " where ftype_code = " & vFtype_code
            Else
                gconDMIS.Execute "insert into HRMS_Header " & _
                                 "(ftype_code,tin,branch_code,retrn_period)" & _
                               " values (" & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ")"
            End If
            Close #1
        Else
            MsgSpeechBox "Process was Cancelled!"
        End If
    Else
        ShowNoRecord
    End If
End Sub

Sub WITH_PREV()
    Dim YY                                                            As String
    YY = cboyear.Text
    Dim vSchedule_num, vFtype_code, vTin_Empyr, vBranch_Code_Emplyr, vRetrn_Period As String
    Dim vSeq_num                                                      As Integer
    Dim vTin, vBranch_Code, vLast_Name, vFirst_Name, vMiddle_Name     As String
    Dim vPrev_Nontax_13th_month, vPrev_Nontax_SSS_Etc, vPrev_Nontax_Salaries As Double
    Dim vPrev_Taxable_13th_month, vPrev_Taxable_Salaries, vPres_Nontax_13th_month As Double
    Dim vPres_Nontax_SSS_Etc, vPres_Nontax_Salaries, vPres_Taxable_13th_month As Double
    Dim vPres_Taxable_Salaries, vExmpn_Amt, vPremium_Paid             As Double
    Dim vTax_Due, vPrev_Tax_wthld, vPres_Tax_wthld                    As Double
    Dim vAmt_Wthld_Dec, vOver_Wthld, vActual_Amt_Wthld                As Double

    Dim vTOTALPrev_Nontax_13th_month, vTOTALPrev_Nontax_SSS_Etc       As Double
    Dim vTOTALPrev_Nontax_Salaries, vTOTALPrev_Taxable_13th_month     As Double
    Dim vTOTALPrev_Taxable_Salaries, vTOTALPres_Nontax_13th_month     As Double
    Dim vTOTALPres_Nontax_SSS_Etc, vTOTALPres_Nontax_Salaries         As Double
    Dim vTOTALPres_Taxable_13th_month, vTOTALPres_Taxable_Salaries    As Double
    Dim vTOTALExmpn_Amt, vTOTALPremium_Paid                           As Double
    Dim vTOTALTax_Due, vTOTALPrev_Tax_wthld                           As Double
    Dim vTOTALPres_Tax_wthld, vTOTALAmt_Wthld_Dec                     As Double
    Dim vTOTALOver_Wthld, vTOTALActual_Amt_Wthld                      As Double

    Dim vCSVSchedule_num, vCSVFtype_code, vCSVTin_Empyr               As String
    Dim vCSVBranch_Code_Emplyr, vCSVRetrn_Period                      As String
    Dim vCSVSeq_num                                                   As Integer
    Dim vCSVTin, vCSVBranch_Code, vCSVLast_Name, vCSVFirst_Name, vCSVMiddle_Name As String
    Dim vCSVPrev_Nontax_13th_month, vCSVPrev_Nontax_SSS_Etc, vCSVPrev_Nontax_Salaries As Double
    Dim vCSVPrev_Taxable_13th_month, vCSVPrev_Taxable_Salaries, vCSVPres_Nontax_13th_month As Double
    Dim vCSVPres_Nontax_SSS_Etc, vCSVPres_Nontax_Salaries, vCSVPres_Taxable_13th_month As Double
    Dim vCSVPres_Taxable_Salaries, vCSVExmpn_Amt, vCSVPremium_Paid    As Double
    Dim vCSVTax_Due, vCSVPrev_Tax_wthld, vCSVPres_Tax_wthld           As Double
    Dim vCSVAmt_Wthld_Dec, vCSVOver_Wthld, vCSVActual_Amt_Wthld       As Double

    Dim I, CNT                                                        As Integer
    Dim schedFName                                                    As String
    Dim Sige                                                          As Boolean
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        I = 1: CNT = 0
        schedFName = EMPLOYER_TIN & ".s74"
        vSchedule_num = "'D7.4'"
        vFtype_code = "'1604CF'"
        vTin_Empyr = "'" & EMPLOYER_TIN & "'"
        vBranch_Code_Emplyr = "NULL"
        vRetrn_Period = "'01/01/" & YY & "'"

        vCSVSchedule_num = "D7.4"
        vCSVFtype_code = "1604CF"
        vCSVTin_Empyr = EMPLOYER_TIN
        vCSVBranch_Code_Emplyr = ""
        vCSVRetrn_Period = "01/01/" & YY
        On Error Resume Next
        Dim MYPATH                                                    As String
        Dim PAYLNAME                                                  As String
        MYPATH = App.Path
        cmdDialogPIS.FILTER = "Schedule Files (*.s74)|*.s74"
        cmdDialogPIS.FilterIndex = 1
        cmdDialogPIS.DefaultExt = "S74"
        PAYLNAME = cmdDialogPIS.Filename
        If MYPATH <> "\" Then
            cmdDialogPIS.Filename = MYPATH & "\" & cmdDialogPIS.Filename
        End If
        If PAYLNAME = "" Then
            cmdDialogPIS.Filename = schedFName
        End If
        cmdDialogPIS.Action = 2
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        FILNAME = cmdDialogPIS.Filename
        If Err = 32755 Then
            dagos = False
        Else
            dagos = True
        End If
        If dagos = True Then
            Open schedFName For Output As #1
            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
            Print #1, "1604CF" & "," & EMPLOYER_TIN & ",," & "01/01/" & YY
            Do While Not rsEmpInfo.EOF
                If Null2String(rsEmpInfo!RESIGNED) = "" Then
                    Sige = True
                Else
                    Sige = False
                    If YEAR(Null2String(rsEmpInfo!RESIGNED)) <> YY Then
                        Sige = True
                    End If
                End If
                If Sige = True Then
                    If Null2String(rsEmpInfo!withprevious) = "Y" Then
                        Set rsYTDDETAILS = New ADODB.Recordset
                        rsYTDDETAILS.Open "select * from HRMS_YTDDetails where YEER = '" & YY & "' AND empno ='" & rsEmpInfo!EMPNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then
                            vSeq_num = I
                            vTin = N2Str2Null(Null2String(rsEmpInfo!tinno))
                            vBranch_Code = N2Str2Null(Null2String(rsEmpInfo!EMPNO))
                            vLast_Name = N2Str2Null(Null2String(rsEmpInfo!lastname))
                            vFirst_Name = N2Str2Null(Null2String(rsEmpInfo!FIRSTNAME))
                            vMiddle_Name = N2Str2Null(Left(Null2String(rsEmpInfo!MIDDLENAME), 1))

                            vPrev_Nontax_13th_month = "0.00"
                            vPrev_Nontax_SSS_Etc = "0.00"
                            vPrev_Nontax_Salaries = "0.00"
                            vPrev_Taxable_13th_month = "0.00"
                            vPrev_Taxable_Salaries = "0.00"

                            vPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)
                            vPres_Taxable_13th_month = "0.00"
                            vPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vPremium_Paid = "0.00"
                            vTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vPrev_Tax_wthld = "0.00"
                            vPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            If vTax_Due > vPres_Tax_wthld Then
                                vAmt_Wthld_Dec = Format(vTax_Due - vPres_Tax_wthld, MAXIMUM_DIGIT)
                            Else
                                vAmt_Wthld_Dec = "0.00"
                            End If
                            If vTax_Due < vPres_Tax_wthld Then
                                vOver_Wthld = Format(vPres_Tax_wthld - vTax_Due, MAXIMUM_DIGIT)
                            Else
                                vOver_Wthld = "0.00"
                            End If
                            If vAmt_Wthld_Dec > vOver_Wthld Then
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld + vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            Else
                                vActual_Amt_Wthld = Format(vPres_Tax_wthld - vOver_Wthld, MAXIMUM_DIGIT)
                            End If

                            'INITIALIZE TOTAL VALUES
                            vTOTALPrev_Nontax_13th_month = vTOTALPrev_Nontax_13th_month + vPrev_Nontax_13th_month
                            vTOTALPrev_Nontax_SSS_Etc = vTOTALPrev_Nontax_SSS_Etc + vPrev_Nontax_SSS_Etc
                            vTOTALPrev_Nontax_Salaries = vTOTALPrev_Nontax_Salaries + vPrev_Nontax_Salaries
                            vTOTALPrev_Taxable_13th_month = vTOTALPrev_Taxable_13th_month + vPrev_Taxable_13th_month
                            vTOTALPrev_Taxable_Salaries = vTOTALPrev_Taxable_Salaries + vPrev_Taxable_Salaries

                            vTOTALPres_Nontax_13th_month = vTOTALPres_Nontax_13th_month + vPres_Nontax_13th_month
                            vTOTALPres_Nontax_SSS_Etc = vTOTALPres_Nontax_SSS_Etc + vPres_Nontax_SSS_Etc
                            vTOTALPres_Nontax_Salaries = vTOTALPres_Nontax_Salaries + vPres_Nontax_Salaries
                            vTOTALPres_Taxable_13th_month = vTOTALPres_Taxable_13th_month + vPres_Taxable_13th_month
                            vTOTALPres_Taxable_Salaries = vTOTALPres_Taxable_Salaries + vPres_Taxable_Salaries
                            vTOTALExmpn_Amt = vTOTALExmpn_Amt + vExmpn_Amt
                            vTOTALPremium_Paid = vTOTALPremium_Paid + vPremium_Paid
                            vTOTALTax_Due = vTOTALTax_Due + vTax_Due
                            vTOTALPrev_Tax_wthld = vTOTALPrev_Tax_wthld + vPrev_Tax_wthld
                            vTOTALPres_Tax_wthld = vTOTALPres_Tax_wthld + vPres_Tax_wthld
                            vTOTALAmt_Wthld_Dec = vTOTALAmt_Wthld_Dec + vAmt_Wthld_Dec
                            vTOTALOver_Wthld = vTOTALOver_Wthld + vOver_Wthld
                            vTOTALActual_Amt_Wthld = vTOTALActual_Amt_Wthld + vActual_Amt_Wthld

                            'INITIALIZE COMMA SEPARATED VALUE
                            vCSVSeq_num = I
                            vCSVTin = Null2String(rsEmpInfo!tinno)
                            vCSVBranch_Code = Null2String(rsEmpInfo!EMPNO)
                            vCSVLast_Name = Null2String(rsEmpInfo!lastname)
                            vCSVFirst_Name = Null2String(rsEmpInfo!FIRSTNAME)
                            vCSVMiddle_Name = Left(Null2String(rsEmpInfo!MIDDLENAME), 1)
                            vCSVPrev_Nontax_13th_month = "0.00"
                            vCSVPrev_Nontax_SSS_Etc = "0.00"
                            vCSVPrev_Nontax_Salaries = "0.00"
                            vCSVPrev_Taxable_13th_month = "0.00"
                            vCSVPrev_Taxable_Salaries = "0.00"
                            vCSVPres_Nontax_13th_month = Format(N2Str2Zero(rsYTDDETAILS!t13thmonth), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_SSS_Etc = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLE), MAXIMUM_DIGIT)
                            vCSVPres_Nontax_Salaries = Format(N2Str2Zero(rsYTDDETAILS!NONTAXABLEADJ), MAXIMUM_DIGIT)
                            vCSVPres_Taxable_13th_month = 0
                            vCSVPres_Taxable_Salaries = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdbasicpay)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commission)) + NumericVal(N2Str2Zero(rsYTDDETAILS!TAXABLEADJ)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remot)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remsal)), MAXIMUM_DIGIT)
                            vCSVExmpn_Amt = Format(N2Str2Zero(rsYTDDETAILS!PersonalEx), MAXIMUM_DIGIT)
                            vCSVPremium_Paid = 0
                            vCSVTax_Due = Format(N2Str2Zero(rsYTDDETAILS!Taxdue), MAXIMUM_DIGIT)
                            vCSVPrev_Tax_wthld = "0.00"
                            vCSVPres_Tax_wthld = Format(NumericVal(N2Str2Zero(rsYTDDETAILS!ytdtax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!commissiontax)) + NumericVal(N2Str2Zero(rsYTDDETAILS!remwtax)), MAXIMUM_DIGIT)
                            vCSVAmt_Wthld_Dec = Format(vAmt_Wthld_Dec, MAXIMUM_DIGIT)
                            vCSVOver_Wthld = Format(vOver_Wthld, MAXIMUM_DIGIT)
                            vCSVActual_Amt_Wthld = Format(vActual_Amt_Wthld, MAXIMUM_DIGIT)

                            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vCSVSeq_num & "," & vCSVTin & "," & vCSVBranch_Code & "," & vCSVLast_Name & "," & vCSVFirst_Name & "," & vCSVMiddle_Name & "," & vCSVPrev_Nontax_13th_month & "," & vCSVPrev_Nontax_SSS_Etc & "," & vCSVPrev_Nontax_Salaries & "," & vCSVPrev_Taxable_13th_month & "," & vCSVPrev_Taxable_Salaries & "," & vCSVPres_Nontax_13th_month & "," & vCSVPres_Nontax_SSS_Etc & "," & vCSVPres_Nontax_Salaries & "," & vCSVPres_Taxable_13th_month & "," & vCSVPres_Taxable_Salaries & "," & vCSVExmpn_Amt & "," & vCSVPremium_Paid & "," & vCSVTax_Due & "," & vCSVPrev_Tax_wthld & "," & vCSVPres_Tax_wthld & "," & vCSVAmt_Wthld_Dec & "," & vCSVOver_Wthld & "," & vCSVActual_Amt_Wthld

                            Set rsDetails74 = New ADODB.Recordset
                            rsDetails74.Open "select * from HRMS_Details74 where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'", gconDMIS
                            If Not rsDetails74.EOF And Not rsDetails74.BOF Then
                                gconDMIS.Execute "update HRMS_Details74 set " & _
                                                 "schedule_num = " & vSchedule_num & ", " & _
                                                 "ftype_code = " & vFtype_code & ", " & _
                                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                                 "retrn_period = " & vRetrn_Period & ", " & _
                                                 "seq_num = " & vSeq_num & ", " & _
                                                 "tin =" & vTin & ", " & _
                                                 "last_name = " & vLast_Name & ", " & _
                                                 "first_name = " & vFirst_Name & ", " & _
                                                 "middle_name = " & vMiddle_Name & ", " & _
                                                 "pres_nontax_13th_month = " & vPres_Nontax_13th_month & ", prev_nontax_13th_month = " & vPrev_Nontax_13th_month & ", " & _
                                                 "pres_nontax_sss_etc = " & vPres_Nontax_SSS_Etc & ", prev_nontax_sss_etc = " & vPrev_Nontax_SSS_Etc & ", " & _
                                                 "pres_nontax_salaries = " & vPres_Nontax_Salaries & ", prev_nontax_salaries = " & vPrev_Nontax_Salaries & ", " & _
                                                 "pres_taxable_13th_month = " & vPres_Taxable_13th_month & ", prev_taxable_13th_month = " & vPrev_Taxable_13th_month & ", " & _
                                                 "pres_taxable_salaries = " & vPres_Taxable_Salaries & ", prev_taxable_salaries = " & vPrev_Taxable_Salaries & ", " & _
                                                 "exmpn_amt = " & vExmpn_Amt & ", " & _
                                                 "premium_paid = " & vPremium_Paid & ", " & _
                                                 "tax_due = " & vTax_Due & ", " & _
                                                 "prev_tax_wthld = " & vPrev_Tax_wthld & ", " & _
                                                 "pres_tax_wthld = " & vPres_Tax_wthld & ", " & _
                                                 "amt_wthld_dec = " & vAmt_Wthld_Dec & ", over_wthld = " & vOver_Wthld & ", actual_amt_wthld = " & vActual_Amt_Wthld & _
                                               " where branch_code = '" & Null2String(rsEmpInfo!EMPNO) & "'"
                            Else
                                gconDMIS.Execute "insert into HRMS_Details74 " & _
                                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,seq_num,tin,branch_code,last_name,first_name,middle_name,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,prev_nontax_13th_month,prev_nontax_sss_etc,prev_nontax_salaries,prev_taxable_13th_month,prev_taxable_salaries,exmpn_amt,premium_paid,tax_due,prev_tax_wthld,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ", " & vSeq_num & ", " & vTin & ", " & vBranch_Code & ", " & vLast_Name & _
                                                 ", " & vFirst_Name & ", " & vMiddle_Name & _
                                                 ", " & vPres_Nontax_13th_month & ", " & vPres_Nontax_SSS_Etc & _
                                                 ", " & vPres_Nontax_Salaries & ", " & vPres_Taxable_13th_month & ", " & vPres_Taxable_Salaries & _
                                                 ", " & vPrev_Nontax_13th_month & ", " & vPrev_Nontax_SSS_Etc & _
                                                 ", " & vPrev_Nontax_Salaries & ", " & vPrev_Taxable_13th_month & _
                                                 ", " & vPrev_Taxable_Salaries & ", " & vExmpn_Amt & ", " & vPremium_Paid & ", " & vTax_Due & ", " & vPrev_Tax_wthld & ", " & vPres_Tax_wthld & ", " & vAmt_Wthld_Dec & ", " & vOver_Wthld & ", " & vActual_Amt_Wthld & ")"
                            End If
                            'Print #1, "SCHEDULE_NUM,FTYPE_CODE,TIN_EMPYR,BRANCH_CODE_EMPLYR,RETRN_PERIOD,SEQ_NUM,TIN,BRANCH_CODE,LAST_NAME,FIRST_NAME,MIDDLE_NAME,EMPLOYMENT_FROM,EMPLOYMENT_TO,PRES_NONTAX_13TH_MONTH,PRES_NONTAX_SSS_ETC,PRES_NONTAX_SALARIES,PRES_TAXABLE_13TH_MONTH,PRES_TAXABLE_SALARIES,EXMPN_AMT,PREMIUM_PAID,TAX_DUE,PRES_TAX_WTHLD,AMT_WTHLD_DEC,OVER_WTHLD,ACTUAL_AMT_WTHLD"
                            I = I + 1
                        End If
                    End If
                End If
                CNT = CNT + 1
                progYTDProcessing.Value = (CNT / rsEmpInfo.RecordCount) * 100
                labEmpNo.Caption = Int(progYTDProcessing.Value) & "%"
                DoEvents
                rsEmpInfo.MoveNext
            Loop
            Print #1, vCSVSchedule_num & "," & vCSVFtype_code & "," & vCSVTin_Empyr & "," & vCSVBranch_Code_Emplyr & "," & vCSVRetrn_Period & "," & vTOTALPrev_Nontax_13th_month & "," & vTOTALPrev_Nontax_SSS_Etc & "," & vTOTALPrev_Nontax_Salaries & "," & vTOTALPrev_Taxable_13th_month & "," & vTOTALPrev_Taxable_Salaries & "," & vTOTALPres_Nontax_13th_month & "," & vTOTALPres_Nontax_SSS_Etc & "," & vTOTALPres_Nontax_Salaries & "," & vTOTALPres_Taxable_13th_month & "," & vTOTALPres_Taxable_Salaries & "," & vTOTALExmpn_Amt & "," & vTOTALPremium_Paid & "," & vTOTALTax_Due & "," & vTOTALPrev_Tax_wthld & "," & vTOTALPres_Tax_wthld & "," & vTOTALAmt_Wthld_Dec & "," & vTOTALOver_Wthld & "," & vTOTALActual_Amt_Wthld
            Set rsControls74 = New ADODB.Recordset
            rsControls74.Open "select * from HRMS_Controls74 where schedule_num = " & vSchedule_num, gconDMIS
            If Not rsControls74.EOF And Not rsControls74.BOF Then
                gconDMIS.Execute "update HRMS_Controls74 set " & _
                                 "schedule_num = " & vSchedule_num & ", " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin_empyr = " & vTin_Empyr & ", " & _
                                 "branch_code_emplyr = " & vBranch_Code_Emplyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "prev_nontax_13th_month = " & vTOTALPrev_Nontax_13th_month & ", " & _
                                 "prev_nontax_sss_etc = " & vTOTALPrev_Nontax_SSS_Etc & ", " & _
                                 "prev_nontax_salaries = " & vTOTALPrev_Nontax_Salaries & ", " & _
                                 "prev_taxable_13th_month = " & vTOTALPrev_Taxable_13th_month & ", " & _
                                 "prev_taxable_salaries = " & vTOTALPrev_Taxable_Salaries & ", " & _
                                 "pres_nontax_13th_month = " & vTOTALPres_Nontax_13th_month & ", " & _
                                 "pres_nontax_sss_etc = " & vTOTALPres_Nontax_SSS_Etc & ", " & _
                                 "pres_nontax_salaries = " & vTOTALPres_Nontax_Salaries & ", " & _
                                 "pres_taxable_13th_month = " & vTOTALPres_Taxable_13th_month & ", " & _
                                 "pres_taxable_salaries = " & vTOTALPres_Taxable_Salaries & ", " & _
                                 "exmpn_amt = " & vTOTALExmpn_Amt & ", " & _
                                 "premium_paid = " & vTOTALPremium_Paid & ", " & _
                                 "tax_due = " & vTOTALTax_Due & ", " & _
                                 "prev_tax_wthld = " & vTOTALPrev_Tax_wthld & ", " & _
                                 "pres_tax_wthld = " & vTOTALPres_Tax_wthld & ", " & _
                                 "amt_wthld_dec = " & vTOTALAmt_Wthld_Dec & ", over_wthld = " & vTOTALOver_Wthld & ", actual_amt_wthld = " & vTOTALActual_Amt_Wthld & _
                               " where schedule_num = " & vSchedule_num
            Else
                gconDMIS.Execute "insert into HRMS_Controls74 " & _
                                 "(schedule_num,ftype_code,tin_empyr,branch_code_emplyr,retrn_period,pres_nontax_13th_month,pres_nontax_sss_etc,pres_nontax_salaries,pres_taxable_13th_month,pres_taxable_salaries,prev_nontax_13th_month,prev_nontax_sss_etc,prev_nontax_salaries,prev_taxable_13th_month,prev_taxable_salaries,exmpn_amt,premium_paid,tax_due,prev_tax_wthld,pres_tax_wthld,amt_wthld_dec,over_wthld,actual_amt_wthld)" & _
                               " values (" & vSchedule_num & ", " & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & _
                                 ", " & vTOTALPres_Nontax_13th_month & ", " & vTOTALPres_Nontax_SSS_Etc & _
                                 ", " & vTOTALPres_Nontax_Salaries & ", " & vTOTALPres_Taxable_13th_month & ", " & vTOTALPres_Taxable_Salaries & _
                                 ", " & vTOTALPrev_Nontax_13th_month & ", " & vTOTALPrev_Nontax_SSS_Etc & _
                                 ", " & vTOTALPrev_Nontax_Salaries & ", " & vTOTALPrev_Taxable_13th_month & ", " & vTOTALPrev_Taxable_Salaries & _
                                 ", " & vTOTALExmpn_Amt & ", " & vTOTALPremium_Paid & ", " & vTOTALTax_Due & ", " & vTOTALPrev_Tax_wthld & ", " & vTOTALPres_Tax_wthld & ", " & vTOTALAmt_Wthld_Dec & ", " & vTOTALOver_Wthld & ", " & vTOTALActual_Amt_Wthld & ")"
            End If
            Set rsHeader = New ADODB.Recordset
            rsHeader.Open "select * from HRMS_Header where ftype_code = " & vFtype_code, gconDMIS
            If Not rsHeader.EOF And Not rsHeader.BOF Then
                gconDMIS.Execute "update HRMS_Header set " & _
                                 "ftype_code = " & vFtype_code & ", " & _
                                 "tin = " & vTin_Empyr & ", " & _
                                 "retrn_period = " & vRetrn_Period & ", " & _
                                 "branch_code = " & vBranch_Code_Emplyr & _
                               " where ftype_code = " & vFtype_code
            Else
                gconDMIS.Execute "insert into HRMS_Header " & _
                                 "(ftype_code,tin,branch_code,retrn_period)" & _
                               " values (" & vFtype_code & ", " & vTin_Empyr & _
                                 ", " & vBranch_Code_Emplyr & ", " & vRetrn_Period & ")"
            End If
            Close #1
        Else
            MsgSpeechBox "Process was Cancelled!"
        End If
    Else
        ShowNoRecord
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Process", "PROCESS BIR ALPHA-LIST PROCESSING") = False Then Exit Sub

    If cboSelect.Text = "Employees Terminated Before December 31" Then TERMINATED_PROC
    If cboSelect.Text = "Employees as of December 31 with No Previous employer within the year" Then WITH_NO_PREV
    If cboSelect.Text = "Employees as of December 31 with Previous Employer within the year" Then WITH_PREV

    LogAudit "G", "GENERATE BIR PROCESSING", cboSelect

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    EMPLOYER_TIN = COMPANY_TIN
    cboSelect.Clear
    cboSelect.AddItem "Employees Terminated Before December 31"
    cboSelect.AddItem "Employees as of December 31 with No Previous employer within the year"
    cboSelect.AddItem "Employees as of December 31 with Previous Employer within the year"
    'FillcboYear cboyear
    fillcombo_up cboyear
    cboyear.Text = YEAR(LOGDATE)
    labEmpNo.Caption = ""
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

