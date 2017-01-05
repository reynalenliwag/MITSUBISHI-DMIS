VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmSMIS_File_AORRate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AOR/OMA/DI Master File"
   ClientHeight    =   5595
   ClientLeft      =   2145
   ClientTop       =   315
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AORRate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   5370
   Begin Crystal.CrystalReport rptAOR 
      Left            =   255
      Top             =   4275
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3810
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   4710
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "AORRate.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "AORRate.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "AORRate.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "AORRate.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save Prospect"
         Top             =   60
         Width           =   705
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4665
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8229
      Cols            =   5
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   30
   End
End
Attribute VB_Name = "frmSMIS_File_AORRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub initGrid()

    With Grid1
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .BackColorFixed = RGB(90, 158, 214)
        .BackColorFixedSel = RGB(110, 180, 230)
        .BackColorBkg = RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cols = 6
        .Cell(0, 0).Text = ""
        .Column(0).Width = 0



        .Cell(0, 1).Text = "Terms"
        .Column(1).Width = 80
        .Column(1).Alignment = cellCenterCenter
        .Column(1).Mask = cellNumeric
        .Column(1).CellType = cellComboBox

        .ComboBox(1).AddItem ("12")
        .ComboBox(1).AddItem ("18")
        .ComboBox(1).AddItem ("24")
        .ComboBox(1).AddItem ("36")
        .ComboBox(1).AddItem ("48")
        .ComboBox(1).AddItem ("60")
        .ComboBox(1).Font.Name = "Arial"

        .Cell(0, 2).Text = "OMA"
        .Column(2).Width = 60
        .Column(2).Alignment = cellCenterCenter
        .Column(2).Mask = cellValue
        .Column(2).FormatString = "0.00"
        .Column(2).DecimalLength = 2

        .Cell(0, 3).Text = "AOR"
        .Column(3).Alignment = cellCenterCenter
        .Column(3).Width = 60
        .Column(3).Mask = cellValue
        .Column(3).FormatString = "0.00"
        .Column(3).DecimalLength = 2

        .Cell(0, 4).Text = "DI"
        .Column(4).Width = 60
        .Column(4).Alignment = cellRightGeneral
        .Column(4).Mask = cellValue
        .Column(4).FormatString = "0.00"
        .Column(4).DecimalLength = 2


        .Cell(0, 5).Text = "Options"
        .Column(5).Locked = True
        .Column(5).Alignment = cellCenterCenter

    End With
End Sub

Sub FillGrid()
    Dim rsRATE                                                        As ADODB.Recordset
    Dim i                                                             As Long
    Set rsRATE = New ADODB.Recordset
    Call rsRATE.Open("select * from smis_fincom_rate order by id asc", gconDMIS, adOpenKeyset, adLockReadOnly)
    'rpert is oma
    'upert is aor
    'downpayment is DI
    Grid1.Rows = 1
    If rsRATE.RecordCount = 0 Then
        Grid1.AddItem ""
        rsRATE.Close
        Set rsRATE = Nothing
        Exit Sub
    End If

    Do While Not rsRATE.EOF
        i = i + 1
        Grid1.AddItem NumericVal(rsRATE("Term")) & Chr(9) & NumericVal(rsRATE("RPERCT")) & Chr(9) & NumericVal(rsRATE("UPERCT")) & Chr(9) & NumericVal(rsRATE("Downpayment")) & Chr(9) & "DELETE", False
        Grid1.Cell(i, 0).Tag = "U"
        Grid1.Cell(i, 1).Tag = rsRATE("ID").Value

        rsRATE.MoveNext
    Loop
    rsRATE.Close
    Set rsRATE = Nothing

    'Add a blank row
    Grid1.AddItem ""
    'Grid1.TopRow = Grid1.Rows
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveData
    FillGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initGrid
    FillGrid
End Sub

Private Sub Grid1_DblClick()
    If Grid1.Selection.FirstCol = Grid1.Cols - 1 And Grid1.ActiveCell.Text = "DELETE" Then
        If MsgBox("Are You Sure you want to delete this record?", vbInformation + vbYesNo) = vbNo Then Exit Sub
        gconDMIS.Execute ("DELETE FROM SMIS_FINCOM_RATE where ID=" & Grid1.Cell(Grid1.ActiveCell.Row, 1).Tag)
        LogAudit "X", "AOR RATE"
        cmdSave.Enabled = False
        FillGrid
    End If
End Sub

Private Sub Grid1_EditRow(ByVal Row As Long)
    If Row = Grid1.Rows - 1 Then
        Grid1.AddItem ""
        Grid1.Cell(Row, 0).Tag = "N"
        Grid1.SelStart = 1
    Else
        If Grid1.Cell(Row, 0).Tag <> "N" Then
            Grid1.Cell(Row, 0).Tag = "E"
        End If
    End If
    cmdSave.Enabled = True
End Sub

Public Sub SaveData()
    Dim strSQL                                                        As String
    Dim i                                                             As Long
    Dim objRs                                                         As New ADODB.Recordset

    Screen.MousePointer = 11
    Dim vTerm, vRPerct, vUPerct, vDownPayment


    For i = 1 To Grid1.Rows - 2
        vTerm = Grid1.Cell(i, 1).SingleValue
        vRPerct = Grid1.Cell(i, 2).DoubleValue
        vUPerct = Grid1.Cell(i, 3).DoubleValue
        vDownPayment = Grid1.Cell(i, 4).DoubleValue
        If vTerm <> 0 Then

            Select Case Grid1.Cell(i, 0).Tag
                Case "N"                                     'New records
                    gconDMIS.Execute ("INSERT INTO SMIS_FINCOM_RATE (FINCOMID, Term,RPerct,UPerct,DownPayment) VALUES ( 1, " & vTerm & " ," & vRPerct & "," & vUPerct & "," & vDownPayment & ") ")
                    LogAudit "A", "AOR RATE", "TERM:" & vTerm & " AOR" & vRPerct & "OMA" & vUPerct
                Case "E"                                     'Edited records

                    gconDMIS.Execute ("UPDATE SMIS_FINCOM_RATE SET " _
                                    & " Term=" & vTerm & "," _
                                    & " RPerct=" & vRPerct & "," _
                                    & " UPerct=" & vUPerct & "," _
                                    & " DownPayment=" & vDownPayment & " WHERE ID=" & Grid1.Cell(i, 1).Tag)
                    LogAudit "E", "AOR RATE", "TERM:" & vTerm & " AOR" & vRPerct & "OMA" & vUPerct
            End Select
        End If
    Next i
    For i = 1 To Grid1.Rows - 2
        Grid1.Cell(i, 0).Tag = "U"
    Next i

    Screen.MousePointer = 0
End Sub

