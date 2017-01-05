VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONLTOAdvances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LTO Advances"
   ClientHeight    =   4680
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7860
   ForeColor       =   &H00F5F5F5&
   Icon            =   "LTOAdvances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   7425
      TabIndex        =   5
      Top             =   4100
      Width           =   7455
      Begin VB.TextBox txtTotalSelected 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtTotalAdvances 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5730
         TabIndex        =   6
         Top             =   60
         Width           =   1635
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label56 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label57 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5100
         TabIndex        =   9
         Top             =   90
         Width           =   435
      End
      Begin VB.Label Label58 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5595
         TabIndex        =   8
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.PictureBox picPettyCashExpenses 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   60
      ScaleHeight     =   3975
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.CommandButton cmdPrintDet 
         Caption         =   "Print &Detailed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   13
         ToolTipText     =   "Print Detailed"
         Top             =   35
         Width           =   1575
      End
      Begin Crystal.CrystalReport rptPettyCA 
         Left            =   120
         Top             =   3480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "UNLIQUIDATED CASH ADVANCES REPORT"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print &Summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5910
         TabIndex        =   12
         ToolTipText     =   "Print Summary"
         Top             =   35
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   380
         Left            =   1080
         TabIndex        =   1
         Text            =   "Cash Advances"
         Top             =   60
         Width           =   1635
      End
      Begin MSFlexGridLib.MSFlexGrid grdLTOPONDO 
         Height          =   3345
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5900
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Date              |                      Name                         |   Amount           |  T    "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "LTOAdvances.frx":030A
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "LTO Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   80
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label54 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONLTOAdvances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImNothing                                                       As Double

Function SetEmployeeName(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT * FROM CMIS_vw_Vemployee WHERE BOOK = 'I' AND CODE = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetEmployeeName = Null2String(rsSBOOK!DESCNAME)
    End If
    Set rsSBOOK = Nothing
End Function

Sub InitGrid()
    cleargrid grdLTOPONDO
    grdLTOPONDO.FormatString = "  Date              |                      Name                         |   Amount           |  T    "
    grdLTOPONDO.ColWidth(4) = 1
End Sub

Sub StoreMemVars()
    Dim LuvUMaam                                                    As Integer
    Dim HopingULoveMeTooMaam                                        As Double
    Dim UrMyFirstLoveMaam                                           As String
    Dim rsLTOPONDO                                                  As ADODB.Recordset
    Set rsLTOPONDO = New ADODB.Recordset
    Set rsLTOPONDO = gconDMIS.Execute("SELECT * FROM CMIS_LTOPondo WHERE Petty_Code = '002' AND LIQUIDATED = 0 ORDER BY ID ASC")
    If Not rsLTOPONDO.EOF And Not rsLTOPONDO.BOF Then
        rsLTOPONDO.MoveFirst
        InitGrid
        LuvUMaam = 0
        HopingULoveMeTooMaam = 0
        ImNothing = 0
        Do While Not rsLTOPONDO.EOF
            LuvUMaam = LuvUMaam + 1
            If Null2Bool(rsLTOPONDO!Tag) = True Then
                UrMyFirstLoveMaam = "T"
            Else
                UrMyFirstLoveMaam = ""
            End If
            
            grdLTOPONDO.AddItem Null2String(rsLTOPONDO!PETTY_DATE) & Chr(9) & SetEmployeeName(Null2String(rsLTOPONDO!Employee)) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsLTOPONDO!original)) & Chr(9) & UrMyFirstLoveMaam & Chr(9) & rsLTOPONDO!Id
            
            If Null2Bool(rsLTOPONDO!Tag) = True Then
                ImNothing = ImNothing + N2Str2Zero(rsLTOPONDO!original)
                HopingULoveMeTooMaam = HopingULoveMeTooMaam + N2Str2Zero(rsLTOPONDO!original)
            End If
            
            If LuvUMaam = 1 Then
                grdLTOPONDO.RemoveItem 1
            End If
            rsLTOPONDO.MoveNext
        Loop
    End If
    Set rsLTOPONDO = Nothing
    
    txtTotalSelected.Text = ToDoubleNumber(ImNothing)
    txtTotalAdvances.Text = ToDoubleNumber(HopingULoveMeTooMaam)
End Sub

Sub TagLTOPONDO()
    Dim ILoveUMaam                                                  As Variant
    grdLTOPONDO.Col = 4
    If grdLTOPONDO.Text <> "" Then
        ILoveUMaam = grdLTOPONDO.Text
        grdLTOPONDO.Col = 3
        If grdLTOPONDO.Text = "T" Then
            gconDMIS.Execute ("Update CMIS_LTOPondo Set Tag = 0 where id = " & ILoveUMaam)
            grdLTOPONDO.Col = 3
            grdLTOPONDO.Text = ""
            grdLTOPONDO.Col = 2
            ImNothing = ImNothing - NumericVal(grdLTOPONDO.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        Else
            gconDMIS.Execute ("Update CMIS_LTOPondo Set Tag = 1 where id = " & ILoveUMaam)
            grdLTOPONDO.Col = 3
            grdLTOPONDO.Text = "T"
            grdLTOPONDO.Col = 2
            ImNothing = ImNothing + NumericVal(grdLTOPONDO.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptPettyCA.Reset
    rptPettyCA.Formulas(0) = "REPORT_DATE = '" & Format(CURRENT_CUTOFF_DATE, "long date") & "'"
    rptPettyCA.Formulas(1) = "PRINTEDBY=" & N2Str2Null(LOGNAME)
    PrintSQLReport rptPettyCA, CMIS_REPORT_PATH & "PettyLTOCurrentSummary.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "LTO ADVANCES - SUMMARY", Text4
    Exit Sub
    
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrintDet_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptPettyCA.Reset
    rptPettyCA.Formulas(0) = "REPORT_DATE = '" & Format(CURRENT_CUTOFF_DATE, "long date") & "'"
    rptPettyCA.Formulas(1) = "PRINTEDBY=" & N2Str2Null(LOGNAME)
    PrintSQLReport rptPettyCA, CMIS_REPORT_PATH & "PettyLTOCurrent.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "LTO ADVANCES - DETAILED", Text4
    Exit Sub
    
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]" '"." & App.Revision & "]"
    InitGrid
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub grdLTOPONDO_DblClick()
    TagLTOPONDO
End Sub

Private Sub grdLTOPONDO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TagLTOPONDO
End Sub

