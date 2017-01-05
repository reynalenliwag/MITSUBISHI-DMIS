VERSION 5.00
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmSMISSRPCommision 
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14715
      TabIndex        =   1
      Top             =   0
      Width           =   14775
      Begin VB.ComboBox cboModels 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   5460
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5490
         TabIndex        =   3
         Top             =   0
         Width           =   1950
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7515
         TabIndex        =   2
         Top             =   0
         Width           =   1950
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   13679
      BackColorFixed  =   16777215
      BackColorBkg    =   -2147483645
      Cols            =   8
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   50
      ScrollBarStyle  =   0
      FixedCols       =   2
      EnterKeyMoveTo  =   1
   End
End
Attribute VB_Name = "frmSMISSRPCommision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FillCombo "selecT DISTINCT UPPER(MODEL) FROM ALL_MODEL", -1, 0, cboModels
    With Grid1
        .Rows = 31
        .Cell(0, 1).Text = "DESCRIPTION"
        .Column(1).Width = 200
        .Column(1).Locked = True

        .Cell(1, 1).Text = "FSP (SRP)"
        .Cell(2, 1).Text = "Less: Discount"
        .Cell(3, 1).Text = "Cash ESP"
        .Cell(4, 1).Text = "Less: Output Tax"
        .Cell(5, 1).Text = "Net Price"
        .Cell(6, 1).Text = ""
        .Cell(7, 1).Text = ""
        .Cell(8, 1).Text = ""

    End With
End Sub

Private Sub cboModels_Click()
    Dim TEMPRS                         As ADODB.Recordset
    Set TEMPRS = New ADODB.Recordset
    Call TEMPRS.Open("Select Descript from ALL_MODEL where MODEL='" & cboModels.Text & "'", gconDMIS, adOpenKeyset, adLockReadOnly)
    Grid1.Cols = TEMPRS.RecordCount + 2
    Dim fld                            As ADODB.Field
    Dim I                              As Integer

    I = 1
    While Not TEMPRS.EOF
        I = I + 1
        Grid1.Column(I).Width = 150
        Grid1.Column(I).CellType = cellTextBox
        Grid1.Column(I).Mask = cellValue
        Grid1.Column(I).DecimalLength = 2
        Grid1.Cell(0, I).Text = Replace(UCase(TEMPRS.Collect(0)), cboModels.Text, "")
        TEMPRS.MoveNext
    Wend
    Set TEMPRS = Nothing

End Sub
'Private Sub cboModels_Click()
'Dim temprs As ADODB.Recordset
'cboModels.Enabled = False
'
'If cboModels.Text = "ALL" Then
'    Set temprs = gconDMIS.Execute("SELECT ID,CODE, MODEL,DESCRIPT,SRP1, SRP2,SRP3, COMMISSION  FROM   ALL_MODEL ORDER BY MODEL")
'
'Else
'    Set temprs = gconDMIS.Execute("SELECT ID,CODE ,MODEL,DESCRIPT,SRP1, SRP2,SRP3, COMMISSION  FROM   ALL_MODEL where model='" & cboModels.Text & "' ORDER BY 2")
'
'End If
'
''
'    Grid1.Rows = 1
'
'While Not temprs.EOF
'    Grid1.AddItem Null2String(temprs!Id) & Chr(9) _
     '                & Null2String(temprs!code) & Chr(9) _
     '                & temprs!descript & Chr(9) _
     '                & Null2String(temprs!SRP1) & Chr(9) _
     '                & Null2String(temprs!SRP2) & Chr(9) _
     '                & Null2String(temprs!SRP3) & Chr(9) _
     '                & Null2String(temprs!COMMISSION), _
     '                False
'    temprs.MoveNext
'Wend
'Grid1.Refresh
'cboModels.Enabled = True
'End Sub
'
'Private Sub cmdUpdate_Click()
''Dim I As Long
' '   For I = 1 To Grid1.Rows - 1
'  '  SQL = "UPDATE ALL_MODEL SET " & _
   '   '       "  SRP1=" & Grid1.Cell(I, 4).DoubleValue & " ," & _
   '    '      "  SRP2=" & Grid1.Cell(I, 5).DoubleValue & " ," & _
   '     '     "  SRP3=" & Grid1.Cell(I, 6).DoubleValue & " ," & _
   '      '    "  COMMISSION=" & Grid1.Cell(I, 7).DoubleValue & " where id= " & Grid1.Cell(I, 1).Text



'       ' gconDMIS.Execute (SQL)
'    'Next
'    'MessagePop RecSaveInfo, " Record Updated", "Recordupdated"
'End Sub
'
'Private Sub Form_Load()
'    FillCombo "selecT DISTINCT MODEL FROM ALL_MODEL", -1, 0, cboModels
'    cboModels.AddItem ("ALL")
'    cboModels.ListIndex = 0
'
'With Grid1
'.Cell(0, 1).Text = "CODE"
'.Cell(0, 2).Text = "MODEL"
'.Cell(0, 3).Text = "DESCRIPTION"
'.Cell(0, 4).Text = "SRP1"
'.Cell(0, 5).Text = "SRP2"
'.Cell(0, 6).Text = "SRP3"
'.Cell(0, 7).Text = "COMMISION"
'
'.Column(0).Locked = True
'.Column(1).Locked = True
'.Column(2).Locked = True
'.Column(3).Locked = True
'
'
'.Column(1).Width = 0
'.Column(2).Width = 80
'.Column(3).Width = 200
'
'.Column(4).CellType = cellTextBox
'.Column(5).CellType = cellTextBox
'.Column(6).CellType = cellTextBox
'.Column(7).CellType = cellTextBox
'
'.Column(4).DecimalLength = 2
'.Column(5).DecimalLength = 2
'.Column(6).DecimalLength = 2
'.Column(7).DecimalLength = 2
'
'
'
'.Column(4).Mask = cellValue
'.Column(5).Mask = cellValue
'.Column(6).Mask = cellValue
'.Column(7).Mask = cellValue
'
'End With
'
'End Sub
'
'    cboModels.AddItem ("ALL")
'    cboModels.ListIndex = 0
'
'With Grid1
'.Cell(0, 1).Text = "CODE"
'.Cell(0, 2).Text = "MODEL"
'.Cell(0, 3).Text = "DESCRIPTION"
'.Cell(0, 4).Text = "SRP1"
'.Cell(0, 5).Text = "SRP2"
'.Cell(0, 6).Text = "SRP3"
'.Cell(0, 7).Text = "COMMISION"
'
'.Column(0).Locked = True
'.Column(1).Locked = True
'.Column(2).Locked = True
'.Column(3).Locked = True
'
'
'.Column(1).Width = 0
'.Column(2).Width = 80
'.Column(3).Width = 200
'
'.Column(4).CellType = cellTextBox
'.Column(5).CellType = cellTextBox
'.Column(6).CellType = cellTextBox
'.Column(7).CellType = cellTextBox
'
'.Column(4).DecimalLength = 2
'.Column(5).DecimalLength = 2
'.Column(6).DecimalLength = 2
'.Column(7).DecimalLength = 2
'
'
'
'.Column(4).Mask = cellValue
'.Column(5).Mask = cellValue
'.Column(6).Mask = cellValue
'.Column(7).Mask = cellValue
'
'End With

Private Sub Form_Resize()
    Grid1.Width = Me.ScaleWidth
    Grid1.Height = Me.ScaleHeight - Picture1.Height
End Sub
