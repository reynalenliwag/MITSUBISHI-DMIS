VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONPettyCashAdvances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Petty Cash Advances"
   ClientHeight    =   4650
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7635
   ForeColor       =   &H00F5F5F5&
   Icon            =   "PettyCashAdvances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   7350
      TabIndex        =   5
      Top             =   4080
      Width           =   7380
      Begin VB.TextBox txtTotalSelected 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtTotalAdvances 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5650
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
         Left            =   5040
         TabIndex        =   9
         Top             =   90
         Width           =   615
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
         Left            =   5520
         TabIndex        =   8
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.PictureBox picPettyCashExpenses 
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   60
      ScaleHeight     =   3945
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.CommandButton cmdPrintDetailed 
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
         Height          =   435
         Left            =   4320
         TabIndex        =   13
         ToolTipText     =   "Print Detailed"
         Top             =   50
         Width           =   1575
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
         Height          =   435
         Left            =   5910
         TabIndex        =   12
         ToolTipText     =   "Print Summary"
         Top             =   50
         Width           =   1575
      End
      Begin Crystal.CrystalReport rptPettyCA 
         Left            =   120
         Top             =   3090
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
      Begin VB.TextBox Text4 
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
         ForeColor       =   &H00701E2A&
         Height          =   380
         Left            =   1680
         TabIndex        =   1
         Text            =   "Cash Advances"
         Top             =   60
         Width           =   1635
      End
      Begin MSFlexGridLib.MSFlexGrid grdPetty 
         Height          =   3345
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   7395
         _ExtentX        =   13044
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
         MouseIcon       =   "PettyCashAdvances.frx":030A
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   1455
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
         Left            =   1500
         TabIndex        =   3
         Top             =   120
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONPettyCashAdvances"
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
    cleargrid grdPetty
    grdPetty.FormatString = "  Date              |                      Name                         |   Amount           |  T    "
    grdPetty.ColWidth(4) = 1
End Sub

Sub StoreMemVars()
    Dim LuvUMaam                                                    As Integer
    Dim HopingULoveMeTooMaam                                        As Double
    Dim UrMyFirstLoveMaam                                           As String
    Dim rsPETTY                                                     As ADODB.Recordset
    
    Set rsPETTY = New ADODB.Recordset
    'updated nov. 14, 2005
    'Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where PETTY_CODE = '002' AND LIQUIDATED = 0 order by ID asc")
    'updated Aug. 24, 2007
    'Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where PETTY_CODE = '002' AND LIQUIDATED = 0 order by EMPLOYEE asc")
    Set rsPETTY = gconDMIS.Execute("SELECT * FROM CMIS_Petty WHERE (PETTY_DATE <= '" & CASHPOSITION_CUTOFF_DATE & "' AND PETTY_CODE = '002') AND PETTY_CASH > 0 ORDER BY EMPLOYEE ASC")
    If Not rsPETTY.EOF And Not rsPETTY.BOF Then
        rsPETTY.MoveFirst
        InitGrid
        LuvUMaam = 0
        HopingULoveMeTooMaam = 0
        ImNothing = 0
        Do While Not rsPETTY.EOF
            If Null2Date(rsPETTY!LIQ_DATE) < CASHPOSITION_CUTOFF_DATE Then
                LuvUMaam = LuvUMaam + 1
                If Null2Bool(rsPETTY!Tag) = True Then
                    UrMyFirstLoveMaam = "T"
                Else
                    UrMyFirstLoveMaam = ""
                End If
                grdPetty.AddItem Null2String(rsPETTY!PETTY_DATE) & Chr(9) & SetEmployeeName(Null2String(rsPETTY!Employee)) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTY!original)) & Chr(9) & UrMyFirstLoveMaam & Chr(9) & rsPETTY!Id
                
                If Null2Bool(rsPETTY!Tag) = True Then
                    ImNothing = ImNothing + N2Str2Zero(rsPETTY!PETTY_CASH)
                    HopingULoveMeTooMaam = HopingULoveMeTooMaam + N2Str2Zero(rsPETTY!PETTY_CASH)
                End If
                
                If LuvUMaam = 1 Then
                    grdPetty.RemoveItem 1
                End If
            End If
            rsPETTY.MoveNext
        Loop
    End If
    Set rsPETTY = Nothing
    
    txtTotalSelected.Text = ToDoubleNumber(ImNothing)
    txtTotalAdvances.Text = ToDoubleNumber(HopingULoveMeTooMaam)
End Sub

Sub TagPetty()
    Dim ILoveUMaam                                                  As Variant
    
    grdPetty.Col = 4
    If grdPetty.Text <> "" Then
        ILoveUMaam = grdPetty.Text
        grdPetty.Col = 3
        If grdPetty.Text = "T" Then
            gconDMIS.Execute ("UPDATE CMIS_Petty SET tag = 0 WHERE id = " & ILoveUMaam)
            grdPetty.Col = 3
            grdPetty.Text = ""
            grdPetty.Col = 2
            ImNothing = ImNothing - NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        Else
            gconDMIS.Execute ("UPDATE CMIS_Petty SET tag = 1 WHERE id = " & ILoveUMaam)
            grdPetty.Col = 3
            grdPetty.Text = "T"
            grdPetty.Col = 2
            ImNothing = ImNothing + NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptPettyCA.Reset
    rptPettyCA.Formulas(0) = "REPORT_DATE = '" & Format(CURRENT_CUTOFF_DATE, "long date") & "'"

    rptPettyCA.Formulas(3) = "PREPAREDBY='" & PreparedBy & "'"
    rptPettyCA.Formulas(4) = "NOTEDBY='" & ApprovedBy & "'"
    rptPettyCA.Formulas(5) = "CHECKEDBY='" & CheckedBy & "'"

    rptPettyCA.Formulas(6) = "PRINTEDBY=" & N2Str2Null(LOGNAME)

    PrintSQLReport rptPettyCA, CMIS_REPORT_PATH & "PettyCACurrentSummary.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "PETTY CASH ADVANCES - SUMMARY", Text4
End Sub

Private Sub cmdPrintDetailed_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptPettyCA.Reset
    rptPettyCA.Formulas(0) = "REPORT_DATE = '" & Format(CURRENT_CUTOFF_DATE, "long date") & "'"
    rptPettyCA.Formulas(1) = "PRINTEDBY=" & N2Str2Null(LOGNAME)
    PrintSQLReport rptPettyCA, CMIS_REPORT_PATH & "PettyCACurrent.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "PETTY CASH ADVANCES - DETAILED", Text4
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

Private Sub grdPetty_DblClick()
    'TagPetty
End Sub

Private Sub grdPetty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TagPetty
End Sub

