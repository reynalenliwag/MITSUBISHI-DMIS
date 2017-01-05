VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOSMSTransactionInvAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplies Inventory Adjusment"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InventoryAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9375
   Begin VB.PictureBox picSupplyAdjust2 
      Height          =   2835
      Left            =   2700
      ScaleHeight     =   2775
      ScaleWidth      =   3885
      TabIndex        =   2
      Top             =   1110
      Width           =   3945
      Begin VB.TextBox txtCost 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   870
         Width           =   1425
      End
      Begin VB.TextBox txtAdd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtSUPPLY_CODE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   120
         Width           =   1965
      End
      Begin VB.TextBox txtMinus 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1650
         Width           =   1425
      End
      Begin wizButton.cmd cmdSave 
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   2160
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":030A
      End
      Begin wizButton.cmd cmdCancel 
         Height          =   495
         Left            =   1980
         TabIndex        =   14
         Top             =   2160
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":0326
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1890
         TabIndex        =   8
         Top             =   900
         Width           =   645
      End
      Begin VB.Label labSUPPLY_DESCRIPTION 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   510
         Width           =   3615
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost (Add)"
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
         Left            =   150
         TabIndex        =   6
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Add   (+)"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1290
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Code"
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
         Left            =   180
         TabIndex        =   3
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Minus (-)"
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
         Left            =   150
         TabIndex        =   12
         Top             =   1680
         Width           =   1755
      End
   End
   Begin wizButton.cmd cmdSupplyAdjust2 
      Height          =   2955
      Left            =   2640
      TabIndex        =   1
      Top             =   1050
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   5212
      TX              =   "F2 - Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "InventoryAdjustment.frx":0342
   End
   Begin Crystal.CrystalReport rptSupplyAdjustments 
      Left            =   870
      Top             =   4740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Materials Inventory Adjustment Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   420
      Top             =   4740
   End
   Begin VB.PictureBox picSupplyAdjust 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      ScaleHeight     =   465
      ScaleWidth      =   9225
      TabIndex        =   15
      Top             =   5730
      Width           =   9225
      Begin wizButton.cmd cmdAdd 
         Height          =   285
         Left            =   30
         TabIndex        =   16
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         TX              =   "F2 - Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":035E
      End
      Begin wizButton.cmd cmdChange 
         Height          =   285
         Left            =   1860
         TabIndex        =   17
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         TX              =   "F3 - Change"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":037A
      End
      Begin wizButton.cmd cmdDelete 
         Height          =   285
         Left            =   3690
         TabIndex        =   18
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         TX              =   "F4 - Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":0396
      End
      Begin wizButton.cmd cmdPrint 
         Height          =   285
         Left            =   5520
         TabIndex        =   19
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         TX              =   "F5 - Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":03B2
      End
      Begin wizButton.cmd cmdF6 
         Height          =   285
         Left            =   7350
         TabIndex        =   20
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         TX              =   "F6 - Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "InventoryAdjustment.frx":03CE
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdSupplyAdjust 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   9
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOSMSTransactionInvAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplyAdjust As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    cmdSupplyAdjust2.ZOrder 0
    picSupplyAdjust2.ZOrder 0
    initMemvars
    On Error Resume Next
    txtSUPPLY_CODE.SetFocus
End Sub

Private Sub cmdCancel_Click()
    initMemvars
    cmdSupplyAdjust2.ZOrder 1
    picSupplyAdjust2.ZOrder 1
    AddorEdit = ""
End Sub

Private Sub cmdChange_Click()
    grdSupplyAdjust.Col = 0
    If grdSupplyAdjust.Text = "No Entry" Or grdSupplyAdjust.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    AddorEdit = "EDIT"
    cmdSupplyAdjust2.ZOrder 0
    picSupplyAdjust2.ZOrder 0
    initMemvars
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgQuestionBox("Delete Supplies Adjustment Entry, Are you sure?", "Delete a Record") = True Then
        grdSupplyAdjust.Col = 8
        If grdSupplyAdjust.Text <> "" Then
            gconDMIS.Execute "delete from OSMS_Adjustment where id = " & grdSupplyAdjust.Text
            rsRefresh
            initGrid
            FillGrid
        Else
            ShowNothingToDeleteMsg
        End If
    End If
End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    PrintSQLReport rptSupplyAdjustments, OSMS_REPORT_PATH & "SupplyAdjustments.rpt", "", OSMS_DataConn, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim VTXTSUPPLY_CODE As String
    Dim VTXTSUPPLY_DESCRIPTION As String
    Dim VTXTCost As Double
    Dim vtxtAdd, vtxtMinus As Integer
    Dim Vusercode, VLastUpdate, VStatus As String

    VTXTSUPPLY_CODE = UCase(N2Str2Null(txtSUPPLY_CODE.Text))
    VTXTSUPPLY_DESCRIPTION = UCase(N2Str2Null(labSUPPLY_DESCRIPTION.Caption))
    VTXTCost = NumericVal(txtCost.Text)
    vtxtAdd = NumericVal(txtAdd.Text)
    vtxtMinus = NumericVal(txtMinus.Text)
    VStatus = "'N'"
    If vtxtAdd = 0 And vtxtMinus = 0 Then
        MsgSpeechBox "Adjustment must Add or Minus a Quantity!"
        Exit Sub
    End If
    Vusercode = "'" & Left(LOGCODE, 3) & "'"
    VLastUpdate = "'" & LOGDATE & "'"

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into Adjustment " & _
                         "(SUPPLY_CODE,SUPPLY_DESCRIPTION,cost,[add],minus,lastupdate,usercode,status)" & _
                       " values (" & VTXTSUPPLY_CODE & ", " & VTXTSUPPLY_DESCRIPTION & ", " & VTXTCost & ", " & vtxtAdd & ", " & vtxtMinus & _
                         ", " & VLastUpdate & ", " & Vusercode & "," & VStatus & ")"
    Else
        gconDMIS.Execute "UPDATE OSMS_ADJUSTMENT set" & _
                       " SUPPLY_CODE = " & VTXTSUPPLY_CODE & "," & _
                       " SUPPLY_DESCRIPTION = " & VTXTSUPPLY_DESCRIPTION & "," & _
                       " cost = " & VTXTCost & "," & _
                       " [add] = " & vtxtAdd & "," & _
                       " minus = " & vtxtMinus & "," & _
                       " lastupdate = " & VLastUpdate & "," & _
                       " status = " & VStatus & "," & _
                       " usercode = " & Vusercode & _
                       " where id = " & labID.Caption
    End If
    rsRefresh
    cleargrid grdSupplyAdjust
    initGrid
    FillGrid
    initMemvars
    On Error Resume Next
    txtSUPPLY_CODE.SetFocus
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        initMemvars
        cmdSupplyAdjust2.ZOrder 1
        picSupplyAdjust2.ZOrder 1
    Case vbKeyF2
        AddorEdit = "ADD"
        cmdSupplyAdjust2.ZOrder 0
        picSupplyAdjust2.ZOrder 0
        initMemvars
        txtSUPPLY_CODE.Enabled = True
        On Error Resume Next
        txtSUPPLY_CODE.SetFocus
    Case vbKeyF3
        grdSupplyAdjust.Col = 0
        If grdSupplyAdjust.Text = "No Entry" Or grdSupplyAdjust.Text = "TAG NO." Then
            MsgBoxXP "Nothing to Edit!", "Empty Record", XP_OKOnly, msg_Information
            Exit Sub
        End If
        AddorEdit = "EDIT"
        cmdSupplyAdjust2.ZOrder 0
        picSupplyAdjust2.ZOrder 0
        initMemvars
        StoreMemVars
    Case vbKeyF6
        Unload Me
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    cleargrid grdSupplyAdjust
    rsRefresh
    initMemvars
    initGrid
    FillGrid
    cmdSupplyAdjust2.ZOrder 1
    picSupplyAdjust2.ZOrder 1
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsSupplyAdjust = New ADODB.Recordset
    rsSupplyAdjust.Open "Select * from OSMS_Adjustment order by SUPPLY_CODE asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initGrid()
    cleargrid grdSupplyAdjust
    With grdSupplyAdjust
        .Row = 0
        .FormatString = "Supply Code        | Supply Description         | Cost           |   Add     |   Minus   | Last Update | User Code | Status      "
        .ColWidth(8) = 1
    End With
End Sub

Sub FillGrid()
    Dim kcnt As Integer
    kcnt = 0
    Dim VSTATUSTEXT As String
    If Not rsSupplyAdjust.EOF And Not rsSupplyAdjust.BOF Then
        Screen.MousePointer = 11
        rsSupplyAdjust.MoveFirst
        Do While Not rsSupplyAdjust.EOF
            kcnt = kcnt + 1
            If Null2String(rsSupplyAdjust!Status) = "N" Then VSTATUSTEXT = Null2String(rsSupplyAdjust!Status) Else VSTATUSTEXT = "POSTED"
            grdSupplyAdjust.AddItem Null2String(rsSupplyAdjust!Supply_Code) & Chr(9) & _
                                    Null2String(rsSupplyAdjust!Supply_Description) & Chr(9) & _
                                    N2Str2Zero(rsSupplyAdjust!Cost) & Chr(9) & _
                                    N2Str2Zero(rsSupplyAdjust![Add]) & Chr(9) & _
                                    N2Str2Zero(rsSupplyAdjust!minus) & Chr(9) & _
                                    Null2String(rsSupplyAdjust!lastupdate) & Chr(9) & _
                                    Null2String(rsSupplyAdjust!usercode) & Chr(9) & _
                                    VSTATUSTEXT & Chr(9) & _
                                    rsSupplyAdjust!Id
            DoEvents
            If VSTATUSTEXT = "POSTED" Then
                grdSupplyAdjust.Row = kcnt + 1
                grdSupplyAdjust.Col = 7
                grdSupplyAdjust.CellForeColor = vbWhite
                grdSupplyAdjust.CellBackColor = vbRed
            End If
            rsSupplyAdjust.MoveNext
        Loop
        If kcnt <> 0 Then grdSupplyAdjust.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Sub initMemvars()
    txtSUPPLY_CODE.Text = ""
    txtCost.Text = 0
    txtAdd.Text = 0
    txtMinus.Text = 0
    cmdSave.Enabled = False
End Sub

Sub StoreMemVars()
    grdSupplyAdjust.Row = grdSupplyAdjust.Row
    grdSupplyAdjust.Col = 8
    Set rsSupplyAdjust = New ADODB.Recordset
    rsSupplyAdjust.Open "Select * from OSMS_Adjustment where id=" & NumericVal(grdSupplyAdjust.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplyAdjust.EOF And Not rsSupplyAdjust.BOF Then
        labID.Caption = rsSupplyAdjust!Id
        txtSUPPLY_CODE.Text = Null2String(rsSupplyAdjust!Supply_Code)
        labSUPPLY_DESCRIPTION.Caption = Null2String(rsSupplyAdjust!Supply_Description)
        txtCost.Text = N2Str2Zero(rsSupplyAdjust!Cost)
        txtAdd.Text = N2Str2Zero(rsSupplyAdjust![Add])
        txtMinus.Text = N2Str2Zero(rsSupplyAdjust!minus)
        If Null2String(rsSupplyAdjust!Status) = "P" Then
            MsgSpeechBox "Warning: Adjustments in this Supply Code has been Posted!" & vbCrLf & _
                       "         Changes in this Data has been Disabled."
            cmdCancel_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub grdSupplyAdjust_DblClick()
    grdSupplyAdjust.Col = 0
    If grdSupplyAdjust.Text = "No Entry" Or grdSupplyAdjust.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    cmdSupplyAdjust2.ZOrder 0
    picSupplyAdjust2.ZOrder 0
    initMemvars
    StoreMemVars
End Sub

Private Sub txtAdd_Change()
    If NumericVal(txtAdd.Text) > 0 Then
        cmdSave.Enabled = True
        txtMinus.Text = 0
    End If
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtMinus_Change()
    If NumericVal(txtMinus.Text) > 0 Then txtAdd.Text = 0
End Sub

Private Sub txtMinus_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtSUPPLY_CODE_LostFocus()
    If txtSUPPLY_CODE.Text = "" Then Exit Sub
    Dim rsSupplyAdjustDUP As ADODB.Recordset
    Dim rsSupply As ADODB.Recordset
    Set rsSupply = New ADODB.Recordset
    rsSupply.Open "Select onhand,SUPPLY_CODE,SUPPLY_DESCRIPTION,Cost,location from OSMS_SUPPLY where SUPPLY_CODE = " & N2Str2Null(txtSUPPLY_CODE.Text), gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        txtCost.Text = N2Str2Zero(rsSupply!Cost)
        labSUPPLY_DESCRIPTION.Caption = Null2String(rsSupply!Supply_Description)
        cmdSave.Enabled = True
        DoEvents
    Else
        MsgSpeechBox "Error: This Supply Code " & txtSUPPLY_CODE.Text & " doesn't exist in Supplies Master File."
        labSUPPLY_DESCRIPTION.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        txtSUPPLY_CODE.SetFocus
    End If
End Sub
