VERSION 5.00
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmCSMSAddPms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preventive Maintenance Service Schedule"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmAddPms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   11880
   Begin FlexCell.Grid Grid1 
      Height          =   7335
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12938
      BackColor2      =   12648384
      Cols            =   5
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   60
      TabIndex        =   5
      Top             =   -60
      Width           =   11775
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9480
         TabIndex        =   8
         ToolTipText     =   "Refresh"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10530
         TabIndex        =   4
         ToolTipText     =   "Close"
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save/Update PMS Schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7200
         TabIndex        =   3
         ToolTipText     =   "Save/Update PMS Schedule"
         Top             =   240
         Width           =   2265
      End
      Begin VB.CommandButton cmdViewjobs 
         Caption         =   "&View Service Operation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4980
         TabIndex        =   2
         ToolTipText     =   "View Service Operation"
         Top             =   240
         Width           =   2205
      End
      Begin VB.CommandButton cmdAddModel 
         Caption         =   "&Add Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3810
         TabIndex        =   1
         ToolTipText     =   "Add Model"
         Top             =   240
         Width           =   1155
      End
      Begin VB.ComboBox cboModel 
         Height          =   345
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label lablstCode 
         Caption         =   "Label2"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model :"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "<<  Insert key - Add row  >>             <<  Delete Key - Erase row  >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   8160
      Width           =   11205
   End
End
Attribute VB_Name = "frmCSMSAddPms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X                                                  As Long

Sub InitGrid()
    With Grid1
        .Cols = 24: .Rows = 2
        .DisplayFocusRect = False: .AllowUserResizing = True
        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)
        .Cell(0, 1).Text = "KM Reading x 1,000"
        .Cell(0, 2).Text = "1"
        .Cell(0, 3).Text = "5"
        .Cell(0, 4).Text = "10"
        .Cell(0, 5).Text = "15"
        .Cell(0, 6).Text = "20"
        .Cell(0, 7).Text = "25"
        .Cell(0, 8).Text = "30"
        .Cell(0, 9).Text = "35"
        .Cell(0, 10).Text = "40"
        .Cell(0, 11).Text = "45"
        .Cell(0, 12).Text = "50"
        .Cell(0, 13).Text = "55"
        .Cell(0, 14).Text = "60"
        .Cell(0, 15).Text = "65"
        .Cell(0, 16).Text = "70"
        .Cell(0, 17).Text = "75"
        .Cell(0, 18).Text = "80"
        .Cell(0, 19).Text = "85"
        .Cell(0, 20).Text = "90"
        .Cell(0, 21).Text = "95"
        .Cell(0, 22).Text = "100"
        .Cell(0, 23).Text = "Code"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox
        .Column(3).CellType = cellTextBox
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox
        .Column(10).CellType = cellTextBox
        .Column(11).CellType = cellTextBox
        .Column(12).CellType = cellTextBox
        .Column(13).CellType = cellTextBox
        .Column(14).CellType = cellTextBox
        .Column(15).CellType = cellTextBox
        .Column(16).CellType = cellTextBox
        .Column(17).CellType = cellTextBox
        .Column(18).CellType = cellTextBox
        .Column(19).CellType = cellTextBox
        .Column(20).CellType = cellTextBox
        .Column(21).CellType = cellTextBox
        .Column(22).CellType = cellTextBox
        .Column(23).CellType = cellTextBox

        .Column(0).Width = 15
        .Column(1).Width = 245
        .Column(2).Width = 24
        .Column(3).Width = 24
        .Column(4).Width = 24
        .Column(5).Width = 24
        .Column(6).Width = 24
        .Column(7).Width = 24
        .Column(8).Width = 24
        .Column(9).Width = 24
        .Column(10).Width = 24
        .Column(11).Width = 24
        .Column(12).Width = 24
        .Column(13).Width = 24
        .Column(14).Width = 24
        .Column(15).Width = 24
        .Column(16).Width = 24
        .Column(17).Width = 24
        .Column(18).Width = 24
        .Column(19).Width = 24
        .Column(20).Width = 24
        .Column(21).Width = 24
        .Column(22).Width = 24
        .Column(23).Width = 70: .Column(23).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 18
        .Range(1, 23, .Rows - 1, 23).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Sub CBOload()
    Dim rsAddModel                                     As ADODB.Recordset
    Set rsAddModel = New ADODB.Recordset
    Set rsAddModel = gconDMIS.Execute("select Model,ID from CSMS_PMS_Hd order by Model asc")
    cboModel.Clear
    If Not rsAddModel.EOF And Not rsAddModel.BOF Then
        Do Until rsAddModel.EOF
            cboModel.AddItem rsAddModel![Model]
            rsAddModel.MoveNext
        Loop
    End If
End Sub

Private Sub cboModel_Click()
    cmdViewjobs.Value = True
End Sub

Private Sub cmdSave_Click()
    If Function_Access(LOGID, "Acess_EDIT", "PMS JOBS") = False Then Exit Sub

    On Error GoTo Errorcode
    If cboModel.Text = "" Then
        MsgBox "Model is blank..."
        Exit Sub
    End If
    If MsgBox("Save all entries..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    Dim xMODEL, xPSM_Description, xKM1_1, xKM5_3, xKM10_6, xKM15_9, xKM20_12, xKM25_15, xKM30_18, xKM35_21, xKM40_24, xKM45_27, xKM50_30, xKM55_33, xKM60_36 As String
    Dim xKM65_39, xKM70_42, xKM75_45, xKM80_48, xKM85_51, xKM90_54, xKM95_57, xKM100_60, Xcode As String
    gconDMIS.Execute "delete  from [CSMS_Psm_Det] where [model] = '" & Trim(cboModel) & "'"
    For X = 1 To Grid1.Rows - 1
        xMODEL = N2Str2Null(cboModel.Text)
        xPSM_Description = N2Str2Null(Grid1.Cell(X, 1).Text)
        xKM1_1 = N2Str2Null(Grid1.Cell(X, 2).Text)
        xKM5_3 = N2Str2Null(Grid1.Cell(X, 3).Text)
        xKM10_6 = N2Str2Null(Grid1.Cell(X, 4).Text)
        xKM15_9 = N2Str2Null(Grid1.Cell(X, 5).Text)
        xKM20_12 = N2Str2Null(Grid1.Cell(X, 6).Text)
        xKM25_15 = N2Str2Null(Grid1.Cell(X, 7).Text)
        xKM30_18 = N2Str2Null(Grid1.Cell(X, 8).Text)
        xKM35_21 = N2Str2Null(Grid1.Cell(X, 9).Text)
        xKM40_24 = N2Str2Null(Grid1.Cell(X, 10).Text)
        xKM45_27 = N2Str2Null(Grid1.Cell(X, 11).Text)
        xKM50_30 = N2Str2Null(Grid1.Cell(X, 12).Text)
        xKM55_33 = N2Str2Null(Grid1.Cell(X, 13).Text)
        xKM60_36 = N2Str2Null(Grid1.Cell(X, 14).Text)
        xKM65_39 = N2Str2Null(Grid1.Cell(X, 15).Text)
        xKM70_42 = N2Str2Null(Grid1.Cell(X, 16).Text)
        xKM75_45 = N2Str2Null(Grid1.Cell(X, 17).Text)
        xKM80_48 = N2Str2Null(Grid1.Cell(X, 18).Text)
        xKM85_51 = N2Str2Null(Grid1.Cell(X, 19).Text)
        xKM90_54 = N2Str2Null(Grid1.Cell(X, 20).Text)
        xKM95_57 = N2Str2Null(Grid1.Cell(X, 21).Text)
        xKM100_60 = N2Str2Null(Grid1.Cell(X, 22).Text)
        Xcode = N2Str2Null(Grid1.Cell(X, 23).Text)

        If Grid1.Cell(X, 1).Text <> "" Then
            gconDMIS.Execute "Insert into CSMS_Psm_Det " & _
                           " (model, PSM_Description, KM1_1, KM5_3, KM10_6, KM15_9, KM20_12, KM25_15, KM30_18, KM35_21, KM40_24,KM45_27 , KM50_30, KM55_33, KM60_36, KM65_39, KM70_42, KM75_45, KM80_48, KM85_51, KM90_54, KM95_57, KM100_60, code)" & _
                           " values(" & xMODEL & "," & xPSM_Description & "," & xKM1_1 & "," & xKM5_3 & "," & xKM10_6 & "," & xKM15_9 & "," & xKM20_12 & "," & xKM25_15 & "," & xKM30_18 & "," & xKM35_21 & "," & xKM40_24 & "," & xKM45_27 & "," & xKM50_30 & "," & xKM55_33 & "," & xKM60_36 & "," & xKM65_39 & "," & xKM70_42 & "," & xKM75_45 & "," & xKM80_48 & "," & xKM85_51 & "," & xKM90_54 & "," & xKM95_57 & "," & xKM100_60 & "," & Xcode & ")"
        End If
    Next X
    LogAudit "A", "PMS JOB DETAIL ", " FOR THE MODEL " & cboModel
    cmdRefresh_Click

    Exit Sub

Errorcode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdAddModel_Click()
    If Function_Access(LOGID, "Acess_ADD", "MODEL") = False Then Exit Sub
    frmCSMSPMSModel.Show 1
    CBOload
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Grid1.Rows = 1
    Dim lng                                            As Integer
    lng = cboModel.ListIndex
    cboModel.ListIndex = -1

    Dim rsLoad                                         As New ADODB.Recordset
    Set rsLoad = New ADODB.Recordset
    Set rsLoad = gconDMIS.Execute("Select code from CSMS_Psm_Det  order by CODE desc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        lablstCode.Caption = Trim(Mid(rsLoad![Code], 2, 10))
        lablstCode.Caption = Format(Val(lablstCode.Caption) + 1, "000000000")
        Grid1.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "P" & Trim(lablstCode.Caption)
    Else
        lablstCode.Caption = "000000000"
        lablstCode.Caption = Format(Val(lablstCode.Caption) + 1, "000000000")
        Grid1.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "P" & Trim(lablstCode.Caption)
    End If
    If lng <> -1 Then
        cboModel.ListIndex = lng
    End If
End Sub

Private Sub cmdViewjobs_Click()
    Dim rsLoad                                         As New ADODB.Recordset
    Set rsLoad = New ADODB.Recordset
    Set rsLoad = gconDMIS.Execute("Select * from CSMS_Psm_Det Where Model = '" & Trim(cboModel.Text) & "' order by CODE asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Grid1.Rows = 1
        Do While Not rsLoad.EOF
            Grid1.AddItem rsLoad![PSM_Description] & vbTab & _
                          rsLoad![KM1_1] & vbTab & _
                          rsLoad![KM5_3] & vbTab & _
                          rsLoad![KM10_6] & vbTab & _
                          rsLoad![KM15_9] & vbTab & _
                          rsLoad![KM20_12] & vbTab & _
                          rsLoad![KM25_15] & vbTab & _
                          rsLoad![KM30_18] & vbTab & _
                          rsLoad![KM35_21] & vbTab & _
                          rsLoad![KM40_24] & vbTab & _
                          rsLoad![KM45_27] & vbTab & _
                          rsLoad![KM50_30] & vbTab & _
                          rsLoad![KM55_33] & vbTab & _
                          rsLoad![KM60_36] & vbTab & _
                          rsLoad![KM65_39] & vbTab & _
                          rsLoad![KM70_42] & vbTab & _
                          rsLoad![KM75_45] & vbTab & _
                          rsLoad![KM80_48] & vbTab & _
                          rsLoad![KM85_51] & vbTab & _
                          rsLoad![KM90_54] & vbTab & _
                          rsLoad![KM95_57] & vbTab & _
                          rsLoad![KM100_60] & vbTab & _
                          rsLoad![Code]
            lablstCode.Caption = Trim(Mid(rsLoad![Code], 2, 10))
            rsLoad.MoveNext
        Loop
    End If

End Sub

Private Sub Form_Activate()
    CenterMe frmMain, Me, 1
    CBOload
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    CBOload
    cmdRefresh_Click
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errorcode:

    If Grid1.Rows > 2 Then

        If KeyCode = vbKeyDelete Then
            If Function_Access(LOGID, "Acess_DELETE", "PMS JOBS") = False Then Exit Sub
            Grid1.Selection.DeleteByRow
        End If
    End If
    If KeyCode = vbKeyInsert Then

        If Function_Access(LOGID, "Acess_ADD", "PMS JOBS") = False Then Exit Sub
        lablstCode.Caption = Format(Val(lablstCode.Caption) + 1, "000000000")
        Grid1.AddItem vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "P" & Trim(lablstCode.Caption)
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

