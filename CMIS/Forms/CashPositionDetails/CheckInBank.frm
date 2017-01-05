VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCASHPOSITIONCheckInBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check In Bank"
   ClientHeight    =   6180
   ClientLeft      =   180
   ClientTop       =   540
   ClientWidth     =   7755
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CheckInBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   6075
      Left            =   60
      ScaleHeight     =   6075
      ScaleWidth      =   7635
      TabIndex        =   5
      Top             =   60
      Width           =   7635
      Begin VB.TextBox txtTotalCheckDepo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   5760
         TabIndex        =   14
         Top             =   4590
         Width           =   1755
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   4590
         Width           =   1635
      End
      Begin VB.TextBox txtCheckNum 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   4950
         Width           =   1635
      End
      Begin VB.TextBox txtCheckDate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   5310
         Width           =   1635
      End
      Begin VB.TextBox txtDeposit_To 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   5670
         Width           =   5820
      End
      Begin MSFlexGridLib.MSFlexGrid grdBANKDEPO 
         Height          =   4455
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Date              |               Bank Name             |   Time     |  Amount Deposit  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CheckInBank.frx":030A
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3870
         TabIndex        =   16
         Top             =   4620
         Width           =   1695
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Left            =   5640
         TabIndex        =   15
         Top             =   4620
         Width           =   195
      End
      Begin VB.Label Label61 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   4620
         Width           =   195
      End
      Begin VB.Label Label62 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   4620
         Width           =   1515
      End
      Begin VB.Label Label65 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   4980
         Width           =   195
      End
      Begin VB.Label Label66 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   4980
         Width           =   1515
      End
      Begin VB.Label Label67 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   5340
         Width           =   195
      End
      Begin VB.Label Label68 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   5340
         Width           =   1515
      End
      Begin VB.Label Label69 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   5700
         Width           =   195
      End
      Begin VB.Label Label70 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit to Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   5700
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONCheckInBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                  As ADODB.Recordset
    Set rsBankName = New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'B' AND Code = '" & XXX & "'")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankName = rsBankName!DESCNAME
    End If
    Set rsBankName = Nothing
End Function

Sub InitGrid()
    cleargrid grdBANKDEPO
    grdBANKDEPO.FormatString = "Date              |               Bank Name             |   Time     |  Amount Deposit  "
    grdBANKDEPO.ColWidth(4) = 1
End Sub

Sub StoreMemVars()
    Dim TOTAL_CHECK_IN_BANK                                         As Double
    Dim PlsLoveMe                                                   As Integer
    Dim rsBANKDEPO                                                  As ADODB.Recordset
    
    Set rsBANKDEPO = New ADODB.Recordset
    'Set rsBANKDEPO = gconDMIS.Execute("Select * from CMIS_BankDepo WHERE [TYPE] = '2' AND CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "' ORDER BY ID ASC")
    Set rsBANKDEPO = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE [TYPE] = '2' AND DATDEPOSIT = '" & CASHPOSITION_CUTOFF_DATE & "' ORDER BY ID ASC")
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
        rsBANKDEPO.MoveFirst
        PlsLoveMe = 0
        TOTAL_CHECK_IN_BANK = 0
        Do While Not rsBANKDEPO.EOF
            PlsLoveMe = PlsLoveMe + 1
            grdBANKDEPO.AddItem Null2String(rsBANKDEPO!bankcode) & Chr(9) & SetBankName(Null2String(rsBANKDEPO!bankcode)) & Chr(9) & Null2String(rsBANKDEPO!timdeposit) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT)) & Chr(9) & rsBANKDEPO!Id
            TOTAL_CHECK_IN_BANK = TOTAL_CHECK_IN_BANK + N2Str2Zero(rsBANKDEPO!DEPOSIT)
            If PlsLoveMe = 1 Then grdBANKDEPO.RemoveItem 1
            rsBANKDEPO.MoveNext
        Loop
    End If
    txtTotalCheckDepo.Text = ToDoubleNumber(TOTAL_CHECK_IN_BANK)
    Set rsBANKDEPO = Nothing
End Sub

Sub StoreBANKDEPODetails(XXX As Variant)
    Dim rsBANKDEPO2                                                 As ADODB.Recordset
    Set rsBANKDEPO2 = New ADODB.Recordset
    Set rsBANKDEPO2 = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE ID = " & XXX)
    If Not rsBANKDEPO2.EOF And Not rsBANKDEPO2.BOF Then
        txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO2!DEPOSIT))
        txtCheckNum.Text = Null2String(rsBANKDEPO2!CheckNum)
        txtCheckDate.Text = Null2String(rsBANKDEPO2!CheckDate)
        txtDeposit_To.Text = Null2String(rsBANKDEPO2!Deposit_To)
    End If
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

Private Sub grdBANKDEPO_Click()
    grdBANKDEPO.Col = 4
    If grdBANKDEPO.Text <> "" Then
        StoreBANKDEPODetails grdBANKDEPO.Text
    End If
End Sub

Private Sub grdBANKDEPO_GotFocus()
    grdBANKDEPO.Col = 4
    If grdBANKDEPO.Text <> "" Then
        StoreBANKDEPODetails grdBANKDEPO.Text
    End If
End Sub

