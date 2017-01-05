VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCASHPOSITIONLTOExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "L.T.O. Expenses"
   ClientHeight    =   5310
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7770
   ForeColor       =   &H00F5F5F5&
   Icon            =   "LTOExpenses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ReplenishTag 
      Caption         =   "Replenish TAG Expenses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      Picture         =   "LTOExpenses.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Replenish TAG Expenses"
      Top             =   4560
      Width           =   2265
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   7425
      TabIndex        =   5
      Top             =   3970
      Width           =   7455
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
         Left            =   5760
         TabIndex        =   7
         Top             =   60
         Width           =   1610
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
         Left            =   1440
         TabIndex        =   6
         Top             =   60
         Width           =   1635
      End
      Begin VB.Label Label58 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   90
         Width           =   495
      End
      Begin VB.Label Label56 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   90
         Width           =   195
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
         TabIndex        =   8
         Top             =   90
         Width           =   1335
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "LTOExpenses.frx":05E9
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Text            =   "Expenses"
         Top             =   50
         Width           =   1635
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "L.T.O. Type"
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
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label54 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1080
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
    Set rsPETTY = gconDMIS.Execute("SELECT * FROM CMIS_LTOPondo WHERE PETTY_CODE = '001' AND REPLENISH = 0 ORDER BY ID ASC")
    If Not rsPETTY.EOF And Not rsPETTY.BOF Then
        rsPETTY.MoveFirst
        InitGrid
        LuvUMaam = 0
        HopingULoveMeTooMaam = 0
        ImNothing = 0
        Do While Not rsPETTY.EOF
            LuvUMaam = LuvUMaam + 1
            If Null2Bool(rsPETTY!Tag) = True Then
                UrMyFirstLoveMaam = "T"
            Else
                UrMyFirstLoveMaam = ""
            End If
            
            grdPetty.AddItem Null2String(rsPETTY!PETTY_DATE) & Chr(9) & SetEmployeeName(Null2String(rsPETTY!Employee)) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTY!PETTY_CASH)) & Chr(9) & UrMyFirstLoveMaam & Chr(9) & rsPETTY!Id
            
            If Null2Bool(rsPETTY!liquid) = True Then
                If Null2Bool(rsPETTY!Tag) = True Then
                    ImNothing = ImNothing + N2Str2Zero(rsPETTY!PETTY_CASH)
                End If
            Else
                If Null2Bool(rsPETTY!Tag) = True Then
                    ImNothing = ImNothing + N2Str2Zero(rsPETTY!original)
                End If
            End If
            
            HopingULoveMeTooMaam = HopingULoveMeTooMaam + N2Str2Zero(rsPETTY!PETTY_CASH)
            
            If LuvUMaam = 1 Then
                grdPetty.RemoveItem 1
            End If
            rsPETTY.MoveNext
        Loop
    End If
    Set rsPETTY = Nothing
    
    txtTotalSelected.Text = ToDoubleNumber(ImNothing)
    txtTotalExpense.Text = ToDoubleNumber(HopingULoveMeTooMaam)
End Sub

Sub TagPetty()
    Dim ILoveUMaam                                                  As Variant
    grdPetty.Col = 4
    If grdPetty.Text <> "" Then
        ILoveUMaam = grdPetty.Text
        grdPetty.Col = 3
        If grdPetty.Text = "T" Then
            gconDMIS.Execute ("Update CMIS_LTOPondo Set tag = 0 where id = " & ILoveUMaam)
            grdPetty.Col = 3
            grdPetty.Text = ""
            grdPetty.Col = 2
            ImNothing = ImNothing - NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        Else
            gconDMIS.Execute ("Update CMIS_LTOPondo Set tag = 1 where id = " & ILoveUMaam)
            grdPetty.Col = 3
            grdPetty.Text = "T"
            grdPetty.Col = 2
            ImNothing = ImNothing + NumericVal(grdPetty.Text)
            txtTotalSelected.Text = ToDoubleNumber(ImNothing)
        End If
    End If
End Sub

Private Sub ReplenishTag_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Dim GridID                                                      As Long
    Dim Karabs                                                      As Integer
    Dim PettyAmount                                                 As Double
    
    ReplenishTag.Enabled = False
    For Karabs = 1 To grdPetty.Rows - 1
        grdPetty.Row = Karabs
        grdPetty.Col = 4
        If grdPetty.Text <> "" Then
            GridID = grdPetty.Text
            grdPetty.Col = 2
            PettyAmount = NumericVal(grdPetty.Text)
            grdPetty.Col = 3
            If grdPetty.Text = "T" Then
                gconDMIS.Execute ("UPDATE CMIS_LTOPondo SET tag = 0, Replenish = 1 WHERE id = " & GridID)
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " LTO_EXP = LTO_EXP - " & PettyAmount & "," & _
                                  " LTO_REPL = LTO_REPL + " & PettyAmount & "" & _
                                  " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        End If
    Next
    StoreMemVars
    ReplenishTag.Enabled = True
    LogAudit "R", "LTO EXPENSES", Text4
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
    TagPetty
End Sub

Private Sub grdPetty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TagPetty
End Sub

