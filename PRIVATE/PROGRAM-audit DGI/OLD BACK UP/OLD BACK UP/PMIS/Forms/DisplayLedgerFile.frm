VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPMISDisplayLedgerFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Ledger File"
   ClientHeight    =   6510
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11190
   ForeColor       =   &H00DEDFDE&
   Icon            =   "DisplayLedgerFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11190
   Begin VB.PictureBox picTags 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   11190
      TabIndex        =   2
      Top             =   5610
      Width           =   11190
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   9525
         MouseIcon       =   "DisplayLedgerFile.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "DisplayLedgerFile.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   705
      End
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
         Height          =   705
         Left            =   10380
         MouseIcon       =   "DisplayLedgerFile.frx":0775
         MousePointer    =   99  'Custom
         Picture         =   "DisplayLedgerFile.frx":08C7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   705
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Enter starting series of tag number (1,2,3, etc.)"
         Top             =   75
         Width           =   1665
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   90
         Width           =   1545
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdLEDGER 
      Height          =   5445
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      TextStyleFixed  =   3
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "DisplayLedgerFile.frx":0C2D
   End
End
Attribute VB_Name = "frmPMISDisplayLedgerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLEDGER                           As ADODB.Recordset

Private Sub cmdDisplay_Click()
    If Function_Access(LOGID, "Acess_View") = False Then Exit Sub
    cleargrid grdLEDGER
    InitGrid
    rsRefresh
    FillGrid
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    cleargrid grdLEDGER
    InitGrid
    txtPartNo.Text = ""
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsLEDGER = New ADODB.Recordset
    rsLEDGER.Open "Select * from LEDGER where STOCKNO = '" & txtPartNo.Text & "' order by id asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    Dim kim                            As Integer
    With grdLEDGER
        .Row = 0
        .FormatString = "Part No      | Description             | Tran Date.   | Tran #      | Who                         | " & _
                        "Received     | Issued          | Balance        | Unit Cost    | MAC            |  Total Cost     | Status "
    End With
End Sub

Sub FillGrid()
    Dim kcnt                           As Integer
    kcnt = 0
    If Not rsLEDGER.EOF And Not rsLEDGER.BOF Then
        Screen.MousePointer = 11
        rsLEDGER.MoveFirst
        Do While Not rsLEDGER.EOF
            kcnt = kcnt + 1
            grdLEDGER.AddItem Null2String(rsLEDGER!STOCKNO) & Chr(9) & _
                              Null2String(rsLEDGER!STOCKDESC) & Chr(9) & _
                              Null2String(rsLEDGER!trandate) & Chr(9) & _
                              Null2String(rsLEDGER!Tranno) & Chr(9) & _
                              Null2String(rsLEDGER!Who) & Chr(9) & _
                              Null2String(rsLEDGER!Received) & Chr(9) & _
                              Null2String(rsLEDGER!Issued) & Chr(9) & _
                              Null2String(rsLEDGER!Balance) & Chr(9) & _
                              Null2String(rsLEDGER!UCost) & Chr(9) & _
                              Null2String(rsLEDGER!Mac) & Chr(9) & _
                              Null2String(rsLEDGER!TtlCost) & Chr(9) & _
                              Null2String(rsLEDGER!Status)
            rsLEDGER.MoveNext
        Loop
        If kcnt <> 0 Then grdLEDGER.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtPARTNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdDisplay_Click
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub
