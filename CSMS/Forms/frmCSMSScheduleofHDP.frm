VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSScheduleMPR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedules of MPR "
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmCSMSScheduleofHDP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   4560
   Begin VB.PictureBox picPROG 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   90
      ScaleHeight     =   1095
      ScaleWidth      =   4365
      TabIndex        =   20
      Top             =   3660
      Visible         =   0   'False
      Width           =   4395
      Begin MSComctlLib.ProgressBar prb 
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   330
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   4365
         _Version        =   655364
         _ExtentX        =   7699
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "COMPUTING REPAIR ORDER DETAILS..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
      Begin VB.Label lblCAP 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   22
         Top             =   690
         Width           =   4155
      End
   End
   Begin VB.OptionButton Option9 
      Caption         =   "OTHER BRANDS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   19
      Top             =   3720
      Width           =   2025
   End
   Begin VB.OptionButton optHyundai 
      Caption         =   "HYUNDAI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   18
      Top             =   3720
      Value           =   -1  'True
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      FillStyle       =   0  'Solid
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
      Height          =   3555
      Left            =   90
      ScaleHeight     =   3525
      ScaleWidth      =   4365
      TabIndex        =   8
      Top             =   60
      Width           =   4395
      Begin VB.OptionButton Option8 
         Caption         =   "Units Received for the Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   1419
         Width           =   4545
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Unit Serviced For The Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   16
         Top             =   3000
         Width           =   4395
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Schedule Of Productivity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   15
         Top             =   2628
         Width           =   3585
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Vehicle Sales Released Last Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   150
         TabIndex        =   14
         Top             =   2190
         Width           =   3795
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Vehicle Sales Released Last 3 Months"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   13
         Top             =   1815
         Width           =   4065
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Units Released for the Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   12
         Top             =   1046
         Width           =   4545
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Schedule Of Service Personnel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   11
         Top             =   673
         Width           =   4365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Schedules Of Total Workshop Sales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   3975
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   -270
         TabIndex        =   10
         Top             =   -30
         Width           =   4785
         _Version        =   655364
         _ExtentX        =   8440
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Choose Schedule"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
   End
   Begin VB.ComboBox cboType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmCSMSScheduleofHDP.frx":1082
      Left            =   5580
      List            =   "frmCSMSScheduleofHDP.frx":1092
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select month from the list"
      Top             =   180
      Width           =   2745
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "frmCSMSScheduleofHDP.frx":10D8
      Left            =   135
      List            =   "frmCSMSScheduleofHDP.frx":10DA
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   4395
      Width           =   2295
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2625
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   4395
      Width           =   1845
   End
   Begin Crystal.CrystalReport rptWorkSales 
      Left            =   180
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Hyundai Dealer Monthly Performance Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2280
      MouseIcon       =   "frmCSMSScheduleofHDP.frx":10DC
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSScheduleofHDP.frx":122E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   4860
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1560
      MouseIcon       =   "frmCSMSScheduleofHDP.frx":1679
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSScheduleofHDP.frx":17CB
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   4860
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
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
      Left            =   5610
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   15
      TabIndex        =   5
      Top             =   4125
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2355
      TabIndex        =   4
      Top             =   4125
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSScheduleMPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function CHECKIFHYUNDAI(PLATE_NO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT MAKE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & PLATE_NO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If UCase(Null2String(rstmp!Make)) = "HYUNDAI" Then
            CHECKIFHYUNDAI = True
        ElseIf Null2String(rstmp!Make) = "" Then
            CHECKIFHYUNDAI = False
        Else
            CHECKIFHYUNDAI = False
        End If
    Else
        CHECKIFHYUNDAI = False
    End If
    Set rstmp = Nothing
End Function

Sub GetBPandGJCustomerInsuranceShare()
    Dim rsREPOR                                        As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim INS_LABOR_TMP                                  As Double
    Dim INS_PART_TMP                                   As Double
    Dim INS_MAT_TMP                                    As Double
    Dim MPR_AMOUNT                                     As Double

    picPROG.Visible = True
    Dim X                                              As Integer
    Dim cnt                                            As Integer
    Dim GJ_INSURANCE As Double: Dim GJ_CUSTOMER        As Double
    Dim BP_INSURANCE As Double: Dim BP_CUSTOMER        As Double
    Dim XDATEX As String
    Dim XROX As String
    
    XROX = "REP_OR = 'R-00003844'"
    XDATEX = "MONTH(DTE_COMP) = 1 AND YEAR(DTE_COMP) = 2009 AND DAY(DTE_COMP) = 12"
    
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE " & XDATEX & "  AND TRANSTYPE = 'R'")
    
    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE MONTH(DTE_COMP) = " & What_month(cboMonth) & " AND YEAR(DTE_COMP) = " & cboYear & " AND TRANSTYPE = 'R'")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        cnt = rsREPOR.RecordCount
        prb.Max = cnt
        prb.Value = 0

        Do While Not rsREPOR.EOF
            DoEvents
            prb.Value = prb.Value + 1
            DoEvents
            INS_LABOR_TMP = NumericVal(rsREPOR!PARTLABOR)
            INS_PART_TMP = NumericVal(rsREPOR!PARTPARTS) + NumericVal(rsREPOR!PARTACCESSORIES)
            INS_MAT_TMP = NumericVal(rsREPOR!PARTMATERIALS)

            Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & Null2String(rsREPOR!REP_OR) & "'")
            If Not (rsDet.BOF And rsDet.EOF) Then
                DoEvents
                lblCap.Caption = "RO NO: " & Null2String(rsREPOR!REP_OR)
                DoEvents

                Do While Not rsDet.EOF
                    GJ_INSURANCE = 0: GJ_CUSTOMER = 0
                    BP_INSURANCE = 0: BP_CUSTOMER = 0

                    MPR_AMOUNT = NumericVal(N2Str2Zero(rsDet!DET_AMT) - N2Str2Zero(rsDet!Discount_2))
                    If rsDet!LIVIL = "1" Then         '----LIVIL 1
                        If Null2String(rsDet!JOBTYPE) <> "BP" And Null2String(rsDet!JOBTYPE) <> "PMS" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_LABOR_TMP > 0 Then
                                    If INS_LABOR_TMP >= MPR_AMOUNT Then
                                        INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT    '
                                        GJ_INSURANCE = GJ_INSURANCE + MPR_AMOUNT    'GJ INSURANCE LABOR
                                    Else
                                        GJ_INSURANCE = GJ_INSURANCE + INS_LABOR_TMP    'GJ INSURANCE LABOR
                                        GJ_CUSTOMER = GJ_CUSTOMER + (MPR_AMOUNT - INS_LABOR_TMP)    'GJ CUSTOMER PART
                                        INS_LABOR_TMP = 0
                                    End If
                                Else
                                    GJ_CUSTOMER = GJ_CUSTOMER + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "PMS" Then
                            If Null2String(rsDet!STATUS1) = "Y" And Null2String(rsDet!wCode) = "W" Then
                                'GJ_I_LABOR = GJ_I_LABOR + MPR_AMOUNT                                                     'GJ INTERNAL LABOR
                            Else
                                If Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    'GJ_I_LABOR = GJ_I_LABOR + MPR_AMOUNT                                                  'GJ INTERNAL LABOR
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    'GJ_W_LABOR = GJ_W_LABOR + MPR_AMOUNT                                                  'GJ WARRANTY LABOR
                                Else
'                                    GJ_INSURANCE = GJ_INSURANCE + INS_LABOR_TMP    'GJ INSURANCE LABOR
'                                    GJ_CUSTOMER = GJ_CUSTOMER + (MPR_AMOUNT - INS_LABOR_TMP)
                                    If INS_LABOR_TMP > 0 Then
                                        If INS_LABOR_TMP >= MPR_AMOUNT Then
                                            INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT    '
                                            GJ_INSURANCE = GJ_INSURANCE + MPR_AMOUNT    'GJ INSURANCE LABOR
                                        Else
                                            GJ_INSURANCE = GJ_INSURANCE + INS_LABOR_TMP    'GJ INSURANCE LABOR
                                            GJ_CUSTOMER = GJ_CUSTOMER + (MPR_AMOUNT - INS_LABOR_TMP)    'GJ CUSTOMER PART
                                            INS_LABOR_TMP = 0
                                        End If
                                    Else
                                        GJ_CUSTOMER = GJ_CUSTOMER + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                    End If
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_LABOR_TMP > 0 Then
                                    If INS_LABOR_TMP >= MPR_AMOUNT Then
                                        INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT
                                        BP_INSURANCE = BP_INSURANCE + MPR_AMOUNT    'BP INSURANCE LABOR
                                    Else
                                        BP_INSURANCE = BP_INSURANCE + INS_LABOR_TMP    'BP INSURANCE LABOR
                                        BP_CUSTOMER = BP_CUSTOMER + (MPR_AMOUNT - INS_LABOR_TMP)    'BP CUSTOMER LABOR
                                        INS_LABOR_TMP = 0
                                    End If
                                Else
                                    BP_CUSTOMER = BP_CUSTOMER + MPR_AMOUNT    'BP CUSTOMER LABOR
                                End If
                            End If
                        End If
                    End If                            '------LIVIL = 1

                    If rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then    '----LIVIL 2 OR 4
                        If Null2String(rsDet!JOBTYPE) <> "BP" And Null2String(rsDet!JOBTYPE) <> "PMS" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_PART_TMP > 0 Then
                                    If INS_PART_TMP >= MPR_AMOUNT Then
                                        INS_PART_TMP = INS_PART_TMP - MPR_AMOUNT    '
                                        GJ_INSURANCE = GJ_INSURANCE + MPR_AMOUNT    'GJ INSURANCE LABOR
                                    Else
                                        GJ_INSURANCE = GJ_INSURANCE + INS_PART_TMP    'GJ INSURANCE LABOR
                                        GJ_CUSTOMER = GJ_CUSTOMER + (MPR_AMOUNT - INS_PART_TMP)    'GJ CUSTOMER PART
                                        INS_PART_TMP = 0
                                    End If
                                Else
                                    GJ_CUSTOMER = GJ_CUSTOMER + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_PART_TMP > 0 Then
                                    If INS_PART_TMP >= MPR_AMOUNT Then
                                        INS_PART_TMP = INS_PART_TMP - MPR_AMOUNT
                                        BP_INSURANCE = BP_INSURANCE + MPR_AMOUNT    'BP INSURANCE LABOR
                                    Else
                                        BP_INSURANCE = BP_INSURANCE + INS_PART_TMP    'BP INSURANCE LABOR
                                        BP_CUSTOMER = BP_CUSTOMER + (MPR_AMOUNT - INS_PART_TMP)    'BP CUSTOMER LABOR
                                        INS_PART_TMP = 0
                                    End If
                                Else
                                    BP_CUSTOMER = BP_CUSTOMER + MPR_AMOUNT    'BP CUSTOMER LABOR
                                End If
                            End If
                        End If
                    End If                            '------LIVIL = 2 OR 4
                    If rsDet!LIVIL = "3" Then         '----LIVIL 3
                        If Null2String(rsDet!JOBTYPE) <> "BP" And Null2String(rsDet!JOBTYPE) <> "PMS" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_MAT_TMP > 0 Then
                                    If INS_MAT_TMP >= MPR_AMOUNT Then
                                        INS_MAT_TMP = INS_MAT_TMP - MPR_AMOUNT    '
                                        GJ_INSURANCE = GJ_INSURANCE + MPR_AMOUNT    'GJ INSURANCE LABOR
                                    Else
                                        GJ_INSURANCE = GJ_INSURANCE + INS_MAT_TMP    'GJ INSURANCE LABOR
                                        GJ_CUSTOMER = GJ_CUSTOMER + (MPR_AMOUNT - INS_MAT_TMP)    'GJ CUSTOMER PART
                                        INS_MAT_TMP = 0
                                    End If
                                Else
                                    GJ_CUSTOMER = GJ_CUSTOMER + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INS_MAT_TMP > 0 Then
                                    If INS_MAT_TMP >= MPR_AMOUNT Then
                                        INS_MAT_TMP = INS_MAT_TMP - MPR_AMOUNT
                                        BP_INSURANCE = BP_INSURANCE + MPR_AMOUNT    'BP INSURANCE LABOR
                                    Else
                                        BP_INSURANCE = BP_INSURANCE + INS_MAT_TMP    'BP INSURANCE LABOR
                                        BP_CUSTOMER = BP_CUSTOMER + (MPR_AMOUNT - INS_MAT_TMP)    'BP CUSTOMER LABOR
                                        INS_MAT_TMP = 0
                                    End If
                                Else
                                    BP_CUSTOMER = BP_CUSTOMER + MPR_AMOUNT    'BP CUSTOMER LABOR
                                End If
                            End If
                        End If
                    End If                            '------LIVIL = 3

                    gconDMIS.Execute ("UPDATE CSMS_RO_DET SET GJ_INSURANCE = " & GJ_INSURANCE & _
                                      ",GJ_CUSTOMER = " & GJ_CUSTOMER & _
                                      ",BP_INSURANCE = " & BP_INSURANCE & _
                                      ",BP_CUSTOMER = " & BP_CUSTOMER & _
                                    " WHERE ID = " & rsDet!ID & "")

                    rsDet.MoveNext
                Loop
            End If
            Set rsDet = Nothing

            'gconDMIS.Execute ("UPDATE CSMS_REPOR SET BP_INSURANCE = " & BP_INSURANCE & _
             '    ",BP_CUSTOMER = " & BP_CUSTOMER & _
             '    ",GJ_CUSTOMER = " & GJ_CUSTOMER & _
             '    ",GJ_INSURANCE = " & GJ_INSURANCE & _
             '    " WHERE REP_OR = '" & Null2String(rsREPOR!REP_OR) & "'")

            rsREPOR.MoveNext
        Loop
    End If
    Set rsREPOR = Nothing
    picPROG.Visible = False
End Sub

Private Sub cmdPrint_Click()
    Dim rsThan                                         As ADODB.Recordset
    Dim MANTH                                          As Integer
    Dim YEER                                           As Integer

    Screen.MousePointer = 11
    If Option1.Value = True Then
        frmMain.Enabled = False
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select dte_comp from CSMS_REPOR where month(dte_comp) = '" & What_month(cboMonth.Text) & "' and year(dte_comp) = " & cboYear.Text & " AND TRANSTYPE = 'R'")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "SCHEDULE OF MONTHLY PERFORMANCE REPORT"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " Year " & cboYear & "'"

            Call GetBPandGJCustomerInsuranceShare

            If optHyundai.Value = True Then
                PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_WorkSales.rpt", "(UCASE({CSMS_CUSVEH.MAKE}) = 'HYUNDAI' OR ISNULL({CSMS_CUSVEH.MAKE}) = TRUE) AND {CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.DTE_COMP}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.DTE_COMP}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
            Else
                PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_WorkSales.rpt", "(UCASE({CSMS_CUSVEH.MAKE}) <> 'HYUNDAI' AND ISNULL({CSMS_CUSVEH.MAKE}) = FALSE) AND {CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.DTE_COMP}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.DTE_COMP}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
            End If
        Else
            ShowNoRecord
        End If
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option1.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        'Call LogAudit("V", Option1.Caption, cboMonth & "-" & cboYear)
        frmMain.Enabled = True
    ElseIf Option2.Value = True Then
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select * from CSMS_SERVICE_PERSONNEL_MAINTENANCE")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Schedules Service Of Service Personnel"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " " & cboYear & "'"

            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_ServicePersonnel.rpt", "", CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option2.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        'Call LogAudit("V", Option2.Caption, cboMonth & "-" & cboYear)

    ElseIf Option3.Value = True Then
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select dte_rel from CSMS_REPOR where month(dte_rel) = '" & What_month(cboMonth.Text) & "' and year(dte_rel) = " & cboYear.Text & "")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Units Released for the Month"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " " & cboYear & "'"
            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_UnitsReleased.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.DTE_REL}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.DTE_REL}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option3.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        'Call LogAudit("V", Option3.Caption, cboMonth & "-" & cboYear)

    ElseIf Option4.Value = True Then
        MANTH = What_month(cboMonth)
        YEER = cboYear
        Dim M1                                         As Integer
        Dim M2                                         As Integer
        Dim YER                                        As String
        Dim X                                          As Date
        Dim Y                                          As Date
        Dim XMONTH                                     As Integer
        Dim XYEAR                                      As Integer

        YER = cboYear
        If What_month(cboMonth) - 2 < 0 Then
            If What_month(cboMonth) = 2 Then
                M1 = 12
            End If
            If What_month(cboMonth) = 1 Then
                M1 = 11
            End If
        Else

        End If
        M2 = What_month(cboMonth)
        Set rsThan = New ADODB.Recordset
        X = firstDay(DateSerial(YEER, MANTH - 3, 1))
        Y = lastDay(DateSerial(YEER, MANTH - 1, 1))

        If What_month(cboMonth) < 4 Then
            If What_month(cboMonth) = 3 Then XMONTH = 12
            If What_month(cboMonth) = 2 Then XMONTH = 11
            If What_month(cboMonth) = 1 Then XMONTH = 10
            XYEAR = cboYear - 1
        Else
            XMONTH = What_month(cboMonth) - 3
            XYEAR = cboYear
        End If

        '"SELECT COUNT(*) AS TOTAL_RELASED FROM SMIS_PURCHAGREE WHERE ISDATE(DATERELEASED) = 1 AND DATERELEASED BETWEEN '" & firstDay(DateSerial(YEER, MANTH - 3, 1)) & "' AND '" & lastDay(DateSerial(YEER, MANTH - 1, 1)) & "'"
        Set rsThan = gconDMIS.Execute("Select DATERELEASED from SMIS_SALESORDER WHERE ISDATE(DATERELEASED) = 1 AND month(DATERELEASED) = " & XMONTH & " and year(DATERELEASED) = " & XYEAR & "")
        'Set rsThan = gconDMIS.Execute("Select DATERELEASED from SMIS_SALESORDER WHERE ISDATE(DATERELEASED) = 1 AND DATERELEASED BETWEEN '" & firstDay(DateSerial(YEER, MANTH - 3, 1)) & "' AND '" & lastDay(DateSerial(YEER, MANTH - 1, 1)) & "'")
        'Set rsThan = gconDMIS.Execute("Select DATERELEASED from SMIS_SALESORDER WHERE DATERELEASED BETWEEN '" & X & "' AND '" & Y & "'")

        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "SCHEDULE OF MONTHLY PERFORMANCE REPORT"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & MonthName(MANTH) & " " & YEER & "'"

            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehiceSalesReleasedLast3Month.rpt", "MONTH({PURCHAGREE.DATERELEASED}) = " & XMONTH & " AND YEAR({PURCHAGREE.DATERELEASED}) = " & XYEAR & "", CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehiceSalesReleasedLast3Month.rpt", "MONTH({PURCHAGREE.DATERELEASED}) >= " & M1 & " AND MONTH({PURCHAGREE.DATERELEASED}) <= " & M2 & " AND YEAR({PURCHAGREE.DATERELEASED}) = " & YER & "", CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehiceSalesReleasedLast3Month.rpt", "MONTH({PURCHAGREE.DATERELEASED}) >= " & M1 & " AND MONTH({PURCHAGREE.DATERELEASED}) <= " & M2 & " AND YEAR({PURCHAGREE.DATERELEASED}) = " & YER & "", CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehiceSalesReleasedLast3Month.rpt", "{PURCHAGREE.DATERELEASED} >= DATE(" & Year(X) & "," & Month(X) & "," & Day(X) & ") AND {PURCHAGREE.DATERELEASED} <= DATE(" & Year(Y) & "," & Month(Y) & "," & Day(Y) & ")", CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehiceSalesReleasedLast3Month.rpt", "{PURCHAGREE.DATERELEASED} between DATE(" & Year(X) & "," & Month(X) & "," & Day(X) & ") AND DATE(" & Year(Y) & "," & Month(Y) & "," & Day(Y) & ")", CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'Call LogAudit("V", Option4.Caption, cboMonth & "-" & cboYear)
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option4.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

    ElseIf Option5.Value = True Then
        If What_month(cboMonth) = 1 Then
            MANTH = 12
            YEER = cboYear - 1
        Else
            MANTH = What_month(cboMonth) - 1
            YEER = cboYear
        End If

        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select DATERELEASED from SMIS_SALESORDER where month(DATERELEASED) = " & MANTH & " and year(DATERELEASED) = " & YEER & "")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Vehicle Sales Released Last Month"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & MonthName(MANTH) & " " & YEER & "'"

            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehicleSalesReleasedLastMonth.rpt", "month({purchagree.datereleased}) >= " & MANTH & " AND month({purchagree.datereleased}) <= " & MANTH & " AND year({purchagree.datereleased}) = " & YEER, DMIS_REPORT_Connection, 1
            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_VehicleSalesReleasedLastMonth.rpt", "MONTH({PURCHAGREE.DATERELEASED}) = " & MANTH & " AND YEAR({PURCHAGREE.DATERELEASED}) = " & YEER & "", CSMS_REPORT_CONNECTION, 1

        Else
            Call ShowNoRecord
        End If
        'Call LogAudit("V", Option5.Caption, cboMonth & "-" & cboYear)
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option5.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

    ElseIf Option6.Value = True Then
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select dte_comp from CSMS_REPOR where month(dte_comp) = '" & What_month(cboMonth.Text) & "' and year(dte_comp) = " & cboYear.Text & "")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Schedule of Productivity Report"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " " & cboYear & "'"

            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_Productivity.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.dte_comp}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.dte_comp}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_Productivity.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.DTE_COMP}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.DTE_COMP}) = " & cboYear.Text & " and day({CSMS_REPOR.DTE_COMP}) = " & 2, CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'Call LogAudit("V", Option6.Caption, cboMonth & "-" & cboYear)
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option6.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

    ElseIf Option7.Value = True Then
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select dte_recd from CSMS_REPOR where month(dte_comp) = '" & What_month(cboMonth.Text) & "' and year(dte_comp) = " & cboYear.Text & "")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Schedule of Total Units Serviced"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " " & cboYear & "'"

            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_TotalUnitsServiced.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.dte_comp}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.dte_comp}) = " & cboYear.Text & " AND {CSMS_REPOR.INVOICE} <> 'PDI RO'", CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_TotalUnitsServiced.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.dte_comp}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.dte_comp}) = " & cboYEAR.Text & "  and day({CSMS_REPOR.dte_comp}) = " & 8, CSMS_REPORT_CONNECTION, 1
            'PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_TotalUnitsServiced.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND {CSMS_REPOR.REP_OR} = 'R-00000703'", CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'Call LogAudit("V", Option7.Caption, cboMonth & "-" & cboYear)
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option7.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        Set rsThan = New ADODB.Recordset
        Set rsThan = gconDMIS.Execute("Select dte_recd from CSMS_REPOR where month(dte_recd) = '" & What_month(cboMonth.Text) & "' and year(dte_recd) = " & cboYear.Text & "")
        If Not rsThan.EOF And Not rsThan.BOF Then
            rptWorkSales.WindowTitle = "Schedule of Total Units Received"
            rptWorkSales.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptWorkSales.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            rptWorkSales.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
            rptWorkSales.Formulas(11) = "forthemonth = '" & "For the Month of " & cboMonth & " " & cboYear & "'"

            PrintSQLReport rptWorkSales, CSMS_REPORT_PATH & "MPR_Schedule_CustomerTraffic_UnitsReceived.rpt", "{CSMS_REPOR.TRANSTYPE} = 'R' AND MONTH({CSMS_REPOR.DTE_RECD}) = " & What_month(cboMonth.Text) & " and YEAR({CSMS_REPOR.DTE_RECD}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
        Else
            Call ShowNoRecord
        End If
        'Call LogAudit("V", Option8.Caption, cboMonth & "-" & cboYear)
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SCHEDULE MPR", "", "", "", Option8.Caption & "-" & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If

    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SCHEDULE MPR)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SCHEDULE MPR", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    cboType = "WORKSHOP SALES": cboType.Enabled = False
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        optHyundai.Visible = True
        Option9.Visible = True
    Else
        optHyundai.Visible = False
        Option9.Visible = False
    End If
End Sub

Private Sub Option2_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option3_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option4_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option5_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option6_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option7_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub

Private Sub Option8_Click()
    optHyundai.Visible = False
    Option9.Visible = False
End Sub
