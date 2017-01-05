VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_ServicePersonnel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICE PERSONNEL MAINTENANCE"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_ServicePersonnel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7575
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   60
      ScaleHeight     =   915
      ScaleWidth      =   7425
      TabIndex        =   44
      Top             =   4050
      Width           =   7455
      Begin VB.TextBox txtBodyAndPaint 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   6090
         TabIndex        =   51
         Top             =   330
         Width           =   945
      End
      Begin VB.TextBox txtGeneralJob 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   3540
         TabIndex        =   49
         Top             =   330
         Width           =   945
      End
      Begin VB.TextBox txtQuickService 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1290
         TabIndex        =   47
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body and Paint"
         Height          =   225
         Index           =   21
         Left            =   4830
         TabIndex        =   50
         Top             =   390
         Width           =   1230
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Job"
         Height          =   225
         Index           =   20
         Left            =   2580
         TabIndex        =   48
         Top             =   390
         Width           =   930
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Service"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   46
         Top             =   390
         Width           =   1095
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   7935
         _Version        =   655364
         _ExtentX        =   13996
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "SERVICE CAPACITY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   3810
      ScaleHeight     =   3945
      ScaleWidth      =   3645
      TabIndex        =   2
      Top             =   30
      Width           =   3675
      Begin VB.TextBox txtSAOther 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   43
         Top             =   3240
         Width           =   945
      End
      Begin VB.TextBox txtSABill 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   41
         Top             =   2850
         Width           =   945
      End
      Begin VB.TextBox txtSAIH 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   40
         Top             =   2460
         Width           =   945
      End
      Begin VB.TextBox txtSAWarr 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   39
         Top             =   2100
         Width           =   945
      End
      Begin VB.TextBox txtSAForeMan 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   38
         Top             =   1710
         Width           =   945
      End
      Begin VB.TextBox txtSANew 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   37
         Top             =   1320
         Width           =   945
      End
      Begin VB.TextBox txtSACert 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   36
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox txtSAMaster 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   35
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
         Height          =   225
         Index           =   12
         Left            =   1710
         TabIndex        =   42
         Top             =   3330
         Width           =   540
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Billing Staff"
         Height          =   225
         Index           =   19
         Left            =   1320
         TabIndex        =   34
         Top             =   2940
         Width           =   930
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In House Instructor"
         Height          =   225
         Index           =   18
         Left            =   690
         TabIndex        =   33
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty"
         Height          =   225
         Index           =   17
         Left            =   1470
         TabIndex        =   32
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ForeMan"
         Height          =   225
         Index           =   16
         Left            =   1500
         TabIndex        =   31
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         Height          =   225
         Index           =   15
         Left            =   1920
         TabIndex        =   30
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certified"
         Height          =   225
         Index           =   13
         Left            =   1605
         TabIndex        =   29
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master"
         Height          =   225
         Index           =   11
         Left            =   1710
         TabIndex        =   28
         Top             =   690
         Width           =   570
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Advisor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   10
         Left            =   960
         TabIndex        =   27
         Top             =   300
         Width           =   1290
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   -30
         Width           =   7935
         _Version        =   655364
         _ExtentX        =   13996
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "INDIRECT PERSONNEL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   60
      ScaleHeight     =   3945
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   30
      Width           =   3675
      Begin VB.TextBox txtBPTechCon 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   26
         Top             =   3450
         Width           =   945
      End
      Begin VB.TextBox txtBPTechTin 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   25
         Top             =   3090
         Width           =   945
      End
      Begin VB.TextBox txtBPTechPaint 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   24
         Top             =   2730
         Width           =   945
      End
      Begin VB.TextBox txtGJTechNew 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   23
         Top             =   1650
         Width           =   945
      End
      Begin VB.TextBox txtGJTechCertified 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   22
         Top             =   1290
         Width           =   945
      End
      Begin VB.TextBox txtGJTechExpert 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   21
         Top             =   930
         Width           =   945
      End
      Begin VB.TextBox txtGJTechMaster 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2370
         TabIndex        =   20
         Top             =   570
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contractor Tech'n"
         Height          =   225
         Index           =   9
         Left            =   840
         TabIndex        =   19
         Top             =   3540
         Width           =   1440
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tinsmith"
         Height          =   225
         Index           =   8
         Left            =   1590
         TabIndex        =   18
         Top             =   3210
         Width           =   705
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paint"
         Height          =   225
         Index           =   7
         Left            =   1860
         TabIndex        =   17
         Top             =   2880
         Width           =   405
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In House Tech'n"
         Height          =   225
         Index           =   6
         Left            =   990
         TabIndex        =   16
         Top             =   2520
         Width           =   1320
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         Height          =   225
         Index           =   5
         Left            =   1920
         TabIndex        =   15
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Body And Paint Technician"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   14
         Top             =   2250
         Width           =   2220
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certified"
         Height          =   225
         Index           =   3
         Left            =   1605
         TabIndex        =   13
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expert"
         Height          =   225
         Index           =   2
         Left            =   1770
         TabIndex        =   12
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master"
         Height          =   225
         Index           =   1
         Left            =   1710
         TabIndex        =   11
         Top             =   660
         Width           =   570
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Job Technician"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   10
         Top             =   300
         Width           =   1950
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   -30
         Width           =   7935
         _Version        =   655364
         _ExtentX        =   13996
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "DIRECT PERSONNEL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   -4440
      ScaleHeight     =   960
      ScaleWidth      =   12315
      TabIndex        =   7
      Top             =   5040
      Width           =   12315
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
         Left            =   11220
         MouseIcon       =   "frmCSMS_ServicePersonnel.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ServicePersonnel.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit Window"
         Top             =   75
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   10530
         MouseIcon       =   "frmCSMS_ServicePersonnel.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ServicePersonnel.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Edit Selected Record"
         Top             =   75
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6090
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   4
      Top             =   5040
      Width           =   1590
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
         Height          =   795
         Left            =   690
         MouseIcon       =   "frmCSMS_ServicePersonnel.frx":19E8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ServicePersonnel.frx":1B3A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel"
         Top             =   75
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
         MouseIcon       =   "frmCSMS_ServicePersonnel.frx":1E78
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ServicePersonnel.frx":1FCA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save this Record"
         Top             =   75
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMS_ServicePersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSERVICE                                          As ADODB.Recordset

Sub EnabledPic(COND As Boolean)
    Picture1.Enabled = COND
    Picture2.Enabled = COND
    picAdds.Visible = Not COND
    picSaves.Visible = COND
End Sub

Sub StoreMemvars()
    If Not (rsSERVICE.BOF And rsSERVICE.EOF) Then
        txtGJTechMaster.Text = NumericVal(rsSERVICE!GJ_TECH_MASTER)
        txtGJTechExpert = NumericVal(rsSERVICE!GJ_TECH_EXPERT)
        txtGJTechCertified = NumericVal(rsSERVICE!GJ_TECH_CERTIFIED)
        txtGJTechNew = NumericVal(rsSERVICE!GJ_TECH_NEW)
        txtBPTechPaint = NumericVal(rsSERVICE!BP_TECH_PAINT)
        txtBPTechTin = NumericVal(rsSERVICE!BP_TECH_TINSMITH)
        txtBPTechCon = NumericVal(rsSERVICE!BP_TECH_CONTR)
        txtSAMaster = NumericVal(rsSERVICE!SA_MASTER)
        txtSACert = NumericVal(rsSERVICE!SA_CERTIFIED)
        txtSANew = NumericVal(rsSERVICE!SA_NEW)
        txtSAForeMan = NumericVal(rsSERVICE!ForeMan)
        txtSAWarr = NumericVal(rsSERVICE!WARRANTY)
        txtSAIH = NumericVal(rsSERVICE!IH_INSTRUCTOR)
        txtSABill = NumericVal(rsSERVICE!BILLING_STAFF)
        txtSAOther = NumericVal(rsSERVICE!OTHERS)

        'UPDATED BY: JUN
        'DATE UPDATED: 12-13-2008
        'DESCRIPTION: SERVICE CAPACITY
        txtQuickService = NumericVal(rsSERVICE!SC_QUICKSERVICE)
        txtGeneralJob = NumericVal(rsSERVICE!SC_GENERALJOB)
        txtBodyAndPaint = NumericVal(rsSERVICE!SC_BODYANDPAINT)
    End If
End Sub

Sub rsRefresh()
    Set rsSERVICE = New ADODB.Recordset
    Set rsSERVICE = gconDMIS.Execute("SELECT * FROM [CSMS_SERVICE_PERSONNEL_MAINTENANCE]")
End Sub

Private Sub cmdCancel_Click()
    Call EnabledPic(False)
End Sub

Private Sub cmdEdit_Click()
    Call EnabledPic(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim vGJ_TECH_MASTER                                As Integer
    Dim vGJ_TECH_EXPERT                                As Integer
    Dim vGJ_TECH_CERTIFIED                             As Integer
    Dim vGJ_TECH_NEW                                   As Integer
    Dim vBP_TECH_PAINT                                 As Integer
    Dim vBP_TECH_TINSMITH                              As Integer
    Dim vBP_TECH_CONTR                                 As Integer
    Dim vSA_MASTER                                     As Integer
    Dim vSA_CERTIFIED                                  As Integer
    Dim vSA_NEW                                        As Integer
    Dim vFOREMAN                                       As Integer
    Dim vWARRANTY                                      As Integer
    Dim vIH_INSTRUCTOR                                 As Integer
    Dim vBILLING_STAFF                                 As Integer
    Dim vOTHERS                                        As Integer

    'UPDATED BY: JUN--------------------
    'DATE UPDATED: 12-13-2008
    'DESCRIPTION: SERVICE CAPACITY
    Dim vSC_QUICKSERVICE                               As Integer
    Dim vSC_GENERALJOB                                 As Integer
    Dim vSC_BODYANDPAINT                               As Integer
    'UPDATED BY: JUN--------------------

    vGJ_TECH_MASTER = N2Str2IntZero(txtGJTechMaster)
    vGJ_TECH_EXPERT = N2Str2IntZero(txtGJTechExpert)
    vGJ_TECH_CERTIFIED = N2Str2IntZero(txtGJTechCertified)
    vGJ_TECH_NEW = N2Str2IntZero(txtGJTechNew)
    vBP_TECH_PAINT = N2Str2IntZero(txtBPTechPaint)
    vBP_TECH_TINSMITH = N2Str2IntZero(txtBPTechTin)
    vBP_TECH_CONTR = N2Str2IntZero(txtBPTechCon)
    vSA_MASTER = N2Str2IntZero(txtSAMaster)
    vSA_CERTIFIED = N2Str2IntZero(txtSACert)
    vSA_NEW = N2Str2IntZero(txtSANew)
    vFOREMAN = N2Str2IntZero(txtSAForeMan)
    vWARRANTY = N2Str2IntZero(txtSAWarr)
    vIH_INSTRUCTOR = N2Str2IntZero(txtSAIH)
    vBILLING_STAFF = N2Str2IntZero(txtSABill)
    vOTHERS = N2Str2IntZero(txtSAOther)

    'UPDATED BY: JUN--------------------
    'DATE UPDATED: 12-13-2008
    'DESCRIPTION: SERVICE CAPACITY
    vSC_QUICKSERVICE = N2Str2IntZero(txtQuickService)
    vSC_GENERALJOB = N2Str2IntZero(txtGeneralJob)
    vSC_BODYANDPAINT = N2Str2IntZero(txtBodyAndPaint)
    'UPDATED BY: JUN--------------------



    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_SERVICE_PERSONNEL_MAINTENANCE")
    If (RSTMP.BOF And RSTMP.EOF) Then
        gconDMIS.Execute ("INSERT INTO CSMS_SERVICE_PERSONNEL_MAINTENANCE " & _
                          "(GJ_TECH_MASTER,GJ_TECH_EXPERT,GJ_TECH_CERTIFIED,GJ_TECH_NEW " & _
                          ",BP_TECH_PAINT,BP_TECH_TINSMITH,BP_TECH_CONTR,SA_MASTER,SA_CERTIFIED " & _
                          ",SA_NEW,FOREMAN,WARRANTY,IH_INSTRUCTOR,BILLING_STAFF,OTHERS,SC_QUICKSERVICE,SC_GENERALJOB,SC_BODYANDPAINT) " & _
                          "Values (" & vGJ_TECH_MASTER & "," & vGJ_TECH_EXPERT & "," & vGJ_TECH_CERTIFIED & _
                          "," & vGJ_TECH_NEW & "," & vBP_TECH_PAINT & "," & vBP_TECH_TINSMITH & _
                          "," & vBP_TECH_CONTR & "," & vSA_MASTER & "," & vSA_CERTIFIED & "," & vSA_NEW & _
                          "," & vFOREMAN & "," & vWARRANTY & "," & vIH_INSTRUCTOR & "," & vBILLING_STAFF & _
                          "," & vOTHERS & "," & vSC_QUICKSERVICE & ", " & vSC_GENERALJOB & "," & vSC_BODYANDPAINT & ")")
        LogAudit "A", "SERVER PERSONNEL MAINTENANCE "
        Call ShowSuccessFullyAdded
    Else
        gconDMIS.Execute ("UPDATE CSMS_SERVICE_PERSONNEL_MAINTENANCE " & _
                        " SET GJ_TECH_MASTER = " & vGJ_TECH_MASTER & _
                          ",GJ_TECH_EXPERT = " & vGJ_TECH_EXPERT & _
                          ",GJ_TECH_CERTIFIED = " & vGJ_TECH_CERTIFIED & _
                          ",GJ_TECH_NEW = " & vGJ_TECH_NEW & _
                          ",BP_TECH_PAINT = " & vBP_TECH_PAINT & _
                          ",BP_TECH_TINSMITH = " & vBP_TECH_TINSMITH & _
                          ",BP_TECH_CONTR = " & vBP_TECH_CONTR & _
                          ",SA_MASTER = " & vSA_MASTER & _
                          ",SA_CERTIFIED =" & vSA_CERTIFIED & _
                          ",SA_NEW = " & vSA_NEW & _
                          ",FOREMAN = " & vFOREMAN & _
                          ",WARRANTY = " & vWARRANTY & _
                          ",IH_INSTRUCTOR = " & vIH_INSTRUCTOR & _
                          ",BILLING_STAFF = " & vBILLING_STAFF & _
                          ",OTHERS = " & vOTHERS & _
                          ",SC_QUICKSERVICE = " & vSC_QUICKSERVICE & _
                          ",SC_GENERALJOB = " & vSC_GENERALJOB & _
                          ",SC_BODYANDPAINT = " & vSC_BODYANDPAINT & "")

        LogAudit "E", "SERVER PERSONNEL MAINTENANCE "
        Call ShowSuccessFullyUpdated
    End If

    Call rsRefresh
    Call StoreMemvars
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Call rsRefresh
    Call StoreMemvars
End Sub

Private Sub txtBPTechCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtBPTechPaint_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtBPTechTin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtGJTechCertified_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtGJTechExpert_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtGJTechMaster_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtGJTechNew_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSABill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSACert_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSAForeMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSAIH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSAMaster_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSANew_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSAOther_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSAWarr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

