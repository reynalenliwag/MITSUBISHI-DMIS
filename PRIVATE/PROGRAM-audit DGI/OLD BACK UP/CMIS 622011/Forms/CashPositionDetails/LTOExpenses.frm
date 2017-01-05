VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONLTOExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L.T.O. Expenses"
   ClientHeight    =   4980
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7770
   ForeColor       =   &H00F5F5F5&
   Icon            =   "LTOExpenses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   60
      ScaleHeight     =   465
      ScaleWidth      =   7635
      TabIndex        =   6
      Top             =   4050
      Width           =   7635
      Begin VB.TextBox txtTotalExpense 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtTotalSelected 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   60
         Width           =   1635
      End
      Begin VB.Label Label58 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   5640
         TabIndex        =   12
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label57 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         Height          =   315
         Left            =   4980
         TabIndex        =   11
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label56 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selected"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   90
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replenish TAG Expenses"
      Height          =   315
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Replenish TAG Expenses"
      Top             =   4590
      Width           =   2445
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
      Begin MSFlexGridLib.MSFlexGrid grdPetty 
         Height          =   3345
         Left            =   60
         TabIndex        =   2
         Top             =   480
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "LTOExpenses.frx":030A
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "Expenses"
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "L.T.O. Type"
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label54 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONLTOExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImNothing                                                         As Double

Function SetEmployeeName(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_vw_Vemployee Where BOOK = 'I' and CODE = '" & XXX & "'")
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

Sub StoreMemvars()
    Dim rsPETTY                                                       As ADODB.Recordset
    Set rsPETTY = New ADODB.Recordset
    Set rsPETTY = gconDMIS.Execute("Select * from CMIS_LTOPondo Where PETTY_CODE = '001' AND REPLENISH = 0 order by ID asc")
    Dim LuvUMaam                                                      As Integer
    Dim HopingULoveMeTooMaam                                          As Double
    Dim UrMyFirstLoveMaam                                             As String
    If Not rsPETTY.EOF And Not rsPETTY.BOF Then
        rsPETTY.MoveFirst: InitGrid: LuvUMaam = 0: HopingULoveMeTooMaam = 0: ImNothing = 0
        Do While Not rsPETTY.EOF
            LuvUMaam = LuvUMaam + 1
            If Null2Bool(rsPETTY!Tag) = True Then UrMyFirstLoveMaam = "T" Else UrMyFirstLoveMaam = ""
            grdPetty.AddItem Null2String(rsPETTY!PETTY_DATE) & Chr(9) & SetEmployeeName(Null2String(rsPETTY!Employee)) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTY!PETTY_CASH)) & Chr(9) & UrMyFirstLoveMaam & Chr(9) & rsPETTY!Id
            If Null2Bool(rsPETTY!liquid) = True Then
                If Null2Bool(rsPETTY!Tag) = True Then ImNothing = ImNothing + N2Str2Zero(rsPETTY!PETTY_CASH)
            Else
                If Null2Bool(rsPETTY!Tag) = True Then ImNothing = ImNothing + N2Str2Zero(rsPETTY!original)
            End If
            HopingULoveMeTooMaam = HopingULoveMeTooMaam + N2Str2Zero(rsPETTY!PETTY_CASH)
            If LuvUMaam = 1 Then grdPetty.RemoveItem 1
            rsPETTY.MoveNext
        Loop
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
            gconDMIS.Execute ("update CMIS_LTOPondo Set tag = 0 Where id = " & ILoveUMaam)
            grdPetty.Col = 3: grdPetty.Text = "": grdPetty.Col = 2: ImNothing = ImNothing - NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        Else
            gconDMIS.Execute ("update CMIS_LTOPondo Set tag = 1 Where id = " & ILoveUMaam)
            grdPetty.Col = 3: grdPetty.Text = "T": grdPetty.Col = 2: ImNothing = ImNothing + NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        End If
    End If
End Sub

Private Sub Command1_Click()

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Dim GridID                                                        As Long
    Dim Karabs                                                        As Integer
    Dim PettyAmount                                                   As Double
    Command1.Enabled = False
    For Karabs = 1 To grdPetty.Rows - 1
        grdPetty.Row = Karabs: grdPetty.Col = 4
        If grdPetty.Text <> "" Then
            GridID = grdPetty.Text
            grdPetty.Col = 2: PettyAmount = NumericVal(grdPetty.Text)
            grdPetty.Col = 3
            If grdPetty.Text = "T" Then
                gconDMIS.Execute ("update CMIS_LTOPondo Set tag = 0, Replenish = 1 Where id = " & GridID)
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                " LTO_EXP = LTO_EXP - " & PettyAmount & "," & _
                                " LTO_REPL = LTO_REPL + " & PettyAmount & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        End If
    Next
    StoreMemvars
    Command1.Enabled = True
    LogAudit "R", "LTO EXPENSES", Text4
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub grdPetty_DblClick()
    TagPetty
End Sub

Private Sub grdPetty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TagPetty
End Sub

