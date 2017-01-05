VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Begin VB.Form frmCASHPOSITIONPettyCashExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Petty Cash Expenses"
   ClientHeight    =   5220
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7770
   ForeColor       =   &H00F5F5F5&
   Icon            =   "PettyCashExpenses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   7605
      TabIndex        =   6
      Top             =   3990
      Width           =   7635
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtTotalExpense 
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   5850
         TabIndex        =   7
         Top             =   60
         Width           =   1635
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selected"
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
         Left            =   60
         TabIndex        =   12
         Top             =   90
         Width           =   1170
      End
      Begin VB.Label Label56 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
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
         Left            =   4980
         TabIndex        =   10
         Top             =   90
         Width           =   450
      End
      Begin VB.Label Label58 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replenish TAG Expenses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5475
      Picture         =   "PettyCashExpenses.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Replenish TAG Expenses"
      Top             =   4485
      Width           =   2220
   End
   Begin VB.PictureBox picPettyCashExpenses 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   60
      ScaleHeight     =   3915
      ScaleWidth      =   7605
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.TextBox Text4 
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Expenses"
         Top             =   90
         Width           =   1635
      End
      Begin MSFlexGridLib.MSFlexGrid grdPetty 
         Height          =   3345
         Left            =   60
         TabIndex        =   2
         Top             =   420
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
         MouseIcon       =   "PettyCashExpenses.frx":05E9
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Type"
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONPettyCashExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImNothing                                                         As Double

Function SetEmployeeName(xxx As Variant)
    Dim rsSBOOK                                                       As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_vw_Vemployee Where BOOK = 'I' and CODE = '" & xxx & "'")
    If Not (rsSBOOK.EOF And rsSBOOK.BOF) Then
        SetEmployeeName = Null2String(rsSBOOK!DESCNAME)
    End If
    Set rsSBOOK = Nothing
End Function

Sub InitGrid()
    Call cleargrid(grdPetty)
    grdPetty.FormatString = "  Date              |                      Name                         |   Amount           |  T    "
    grdPetty.ColWidth(4) = 1
End Sub

Sub StoreMemvars()
    Dim rsPETTY                                                       As New ADODB.Recordset
    
    'updated nov. 14, 2005
    'Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where PETTY_CODE = '001' AND REPLENISH = 0 order by ID asc")
    'updated Aug. 24, 2007
    'Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where PETTY_CODE = '001' AND REPLENISH = 0 order by EMPLOYEE asc")
    
    Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where " & _
        " (PETTY_DATE = '" & CASHPOSITION_CUTOFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1') " & _
        " OR " & _
        " (PETTY_DATE < '" & CASHPOSITION_CUTOFF_DATE & "' AND PETTY_CODE = '001' AND REPLENISH <> '1') " & _
        " Order by EMPLOYEE asc")
    Dim LuvUMaam                                                      As Integer
    Dim HopingULoveMeTooMaam                                          As Double
    Dim UrMyFirstLoveMaam                                             As String
    
    If Not (rsPETTY.EOF And rsPETTY.BOF) Then
        rsPETTY.MoveFirst
        Call InitGrid
        LuvUMaam = 0:       HopingULoveMeTooMaam = 0:       ImNothing = 0
        
        Do While Not rsPETTY.EOF
            LuvUMaam = LuvUMaam + 1
            If Null2Bool(rsPETTY!Tag) = True Then
                UrMyFirstLoveMaam = "T"
            Else
                UrMyFirstLoveMaam = ""
            End If
            
            grdPetty.AddItem Null2String(rsPETTY!PETTY_DATE) & _
                Chr(9) & SetEmployeeName(Null2String(rsPETTY!Employee)) & _
                Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTY!PETTY_CASH)) & _
                Chr(9) & UrMyFirstLoveMaam & _
                Chr(9) & rsPETTY!Id
            If Null2Bool(rsPETTY!liquid) = True Then
                If Null2Bool(rsPETTY!Tag) = True Then ImNothing = ImNothing + N2Str2Zero(rsPETTY!PETTY_CASH)
            Else
                If Null2Bool(rsPETTY!Tag) = True Then ImNothing = ImNothing + N2Str2Zero(rsPETTY!original)
            End If
            
            HopingULoveMeTooMaam = HopingULoveMeTooMaam + N2Str2Zero(rsPETTY!PETTY_CASH)
            If LuvUMaam = 1 Then grdPetty.RemoveItem 1
            
            rsPETTY.MoveNext
        Loop
        grdPetty.Refresh
    Else
        Command1.Enabled = False
    End If
    Set rsPETTY = Nothing
    
    txtTotalSelected.Text = ToDoubleNumber(ImNothing)
    txtTotalExpense.Text = ToDoubleNumber(HopingULoveMeTooMaam)
End Sub

Sub TagPetty()
    Dim ILoveUMaam                                                    As Variant
    grdPetty.Col = 4
    
    If grdPetty.Text <> "" Then
        ILoveUMaam = grdPetty.Text: grdPetty.Col = 3
        If grdPetty.Text = "T" Then
            gconDMIS.Execute ("update CMIS_Petty Set tag = 0 Where id = " & ILoveUMaam)
            
            grdPetty.Col = 3: grdPetty.Text = "": grdPetty.Col = 2: ImNothing = ImNothing - NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        Else
            gconDMIS.Execute ("update CMIS_Petty Set tag = 1 Where id = " & ILoveUMaam)
            
            grdPetty.Col = 3: grdPetty.Text = "T": grdPetty.Col = 2: ImNothing = ImNothing + NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        End If
    End If
End Sub

Private Sub Command1_Click()
    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:
    
    If grdPetty.Rows = 2 Then
        If grdPetty.Text = "" Then Exit Sub
    'Else
    '    Exit Sub
    End If
        
    If MsgBox("Replenish this Record, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        
    Dim GridID                                                        As Long
    Dim Karabs                                                        As Integer
    Dim PettyAmount                                                   As Double
    
    For Karabs = 1 To grdPetty.Rows - 1
        grdPetty.Row = Karabs: grdPetty.Col = 4
        If grdPetty.Text <> "" Then
            GridID = grdPetty.Text
            grdPetty.Col = 2:
            PettyAmount = NumericVal(grdPetty.Text)
            grdPetty.Col = 3
            
            If grdPetty.Text = "T" Then
                gconDMIS.Execute ("update CMIS_Petty Set tag = 0 " & _
                    ", Replenish = 1 " & _
                    " Where id = " & GridID)
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                    " EXPENSE = ROUND(EXPENSE,2) - " & PettyAmount & "," & _
                    " REPLENISH = ROUND(REPLENISH,2) + " & PettyAmount & _
                    " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        End If
    Next
    
    Call StoreMemvars
    Command1.Enabled = False
    
    MsgBox "Replenish Finish", vbInformation, "Info."
    
    LogAudit "R", "PETTY CASH EXPENSES", Text4
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    Call InitGrid
    Call StoreMemvars
    
    'If CDate(CASHPOSITION_CUTOFF_DATE) = CDate(LOGDATE) Then
    Command1.Enabled = True
    'Else
    '   Command1.Enabled = False
    'End If
    Screen.MousePointer = 0
End Sub

Private Sub grdPetty_DblClick()
    'If CDate(CASHPOSITION_CUTOFF_DATE) = CDate(LOGDATE) Then
    Call TagPetty
    'Else
    '   MsgBox "Pls Tag only on Current Cut-Off Date", vbInformation, "Not Allowed"
    'End If
End Sub

Private Sub grdPetty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TagPetty
End Sub
