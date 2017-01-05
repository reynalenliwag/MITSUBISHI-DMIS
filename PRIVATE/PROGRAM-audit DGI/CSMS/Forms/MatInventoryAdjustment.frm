VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCSMSMatInvAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Inventory Adjusment"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   ForeColor       =   &H8000000F&
   Icon            =   "MatInventoryAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9375
   Begin VB.PictureBox picMatAdjust 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   5565
      ScaleHeight     =   915
      ScaleWidth      =   5025
      TabIndex        =   15
      Top             =   5775
      Width           =   5025
      Begin VB.CommandButton cmdF6 
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
         Height          =   795
         Left            =   3000
         MouseIcon       =   "MatInventoryAdjustment.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2280
         MouseIcon       =   "MatInventoryAdjustment.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1560
         MouseIcon       =   "MatInventoryAdjustment.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   840
         MouseIcon       =   "MatInventoryAdjustment.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         Picture         =   "MatInventoryAdjustment.frx":1C61
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picMatAdjust2 
      Height          =   3645
      Left            =   2700
      ScaleHeight     =   3585
      ScaleWidth      =   3885
      TabIndex        =   6
      Top             =   1050
      Width           =   3945
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3090
         MouseIcon       =   "MatInventoryAdjustment.frx":1F74
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel Entry"
         Top             =   2775
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2400
         MouseIcon       =   "MatInventoryAdjustment.frx":2404
         MousePointer    =   99  'Custom
         Picture         =   "MatInventoryAdjustment.frx":2556
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   2775
         Width           =   705
      End
      Begin VB.TextBox txtParticular 
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
         Height          =   645
         Left            =   1200
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "MatInventoryAdjustment.frx":28A6
         Top             =   2040
         Width           =   2595
      End
      Begin VB.TextBox txtCost 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   870
         Width           =   1425
      End
      Begin VB.TextBox txtAdd 
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtMatCde 
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
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   120
         Width           =   1965
      End
      Begin VB.TextBox txtMinus 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         TabIndex        =   14
         Top             =   2040
         Width           =   1335
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
         TabIndex        =   12
         Top             =   900
         Width           =   645
      End
      Begin VB.Label labMatDsc 
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1290
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1755
      End
   End
   Begin wizButton.cmd cmdMatAdjust2 
      Height          =   3765
      Left            =   2640
      TabIndex        =   13
      Top             =   990
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   6641
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
      MICON           =   "MatInventoryAdjustment.frx":28AC
   End
   Begin Crystal.CrystalReport rptMatAdjustments 
      Left            =   870
      Top             =   4800
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
      Top             =   4800
   End
   Begin MSFlexGridLib.MSFlexGrid grdMatAdjust 
      Height          =   5730
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   10107
      _Version        =   393216
      Cols            =   10
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmCSMSMatInvAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMatAdjust                                        As ADODB.Recordset
Dim AddorEdit                                          As String

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    cmdMatAdjust2.ZOrder 0
    picMatAdjust2.ZOrder 0
    initMemvars
    On Error Resume Next
    txtMatCde.SetFocus
End Sub

Private Sub cmdCancel_Click()
    initMemvars
    cmdMatAdjust2.ZOrder 1
    picMatAdjust2.ZOrder 1
End Sub

Private Sub cmdChange_Click()
     If Function_Access(LOGID, "Acess_EDIT") = False Then Exit Sub

    grdMatAdjust.Col = 0
    If grdMatAdjust.Text = "No Entry" Or grdMatAdjust.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    AddorEdit = "EDIT"
    cmdMatAdjust2.ZOrder 0
    picMatAdjust2.ZOrder 0
    initMemvars
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE") = False Then Exit Sub
    
    On Error GoTo Errorcode

    If MsgQuestionBox("Delete Materials Adjustment Entry, Are you sure?", "Delete a Record") = True Then
        grdMatAdjust.Col = 9
        If grdMatAdjust.Text <> "" Then
            gconDMIS.Execute "delete from CSMS_MatAdjust where id = " & grdMatAdjust.Text
            rsRefresh
            InitGrid
            FillGrid
        Else
            ShowNothingToDeleteMsg
        End If
    End If
    Exit Sub

Errorcode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT") = False Then Exit Sub
    Screen.MousePointer = 11
    rptMatAdjustments.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMatAdjustments.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptMatAdjustments, CSMS_REPORT_PATH & "MatAdjustments.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim VTXTMatCde                                     As String
    Dim VTXTMatDsc                                     As String
    Dim VTXTCost                                       As Double
    Dim vtxtAdd, vtxtMinus                             As Integer
    Dim Vusercode, VLastUpdate, VStatus, VParticular   As String

    VTXTMatCde = N2Str2Null(txtMatCde.Text)
    VTXTMatDsc = N2Str2Null(labMatDsc.Caption)
    VTXTCost = NumericVal(txtCost.Text)
    vtxtAdd = NumericVal(txtAdd.Text)
    vtxtMinus = NumericVal(txtMinus.Text)
    VStatus = "'N'"
    VParticular = N2Str2Null(txtParticular.Text)
    If vtxtAdd = 0 And vtxtMinus = 0 Then
        MsgSpeechBox "MatAdjustment must Add or Minus a Quantity!"
        Exit Sub
    End If
    Vusercode = "'" & Left(LOGCODE, 3) & "'"
    VLastUpdate = "'" & LOGDATE & "'"

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into CSMS_MatAdjust " & _
                         "(MatCde,MatDsc,cost,[add],minus,lastupdate,usercode,status,Particular)" & _
                       " values (" & VTXTMatCde & ", " & VTXTMatDsc & ", " & VTXTCost & ", " & vtxtAdd & ", " & vtxtMinus & _
                         ", " & VLastUpdate & ", " & Vusercode & "," & VStatus & "," & VParticular & ")"
    Else
        gconDMIS.Execute "update CSMS_MatAdjust set" & _
                       " MatCde = " & VTXTMatCde & "," & _
                       " MatDsc = " & VTXTMatDsc & "," & _
                       " Particular = " & VParticular & "," & _
                       " cost = " & VTXTCost & "," & _
                       " [add] = " & vtxtAdd & "," & _
                       " minus = " & vtxtMinus & "," & _
                       " lastupdate = " & VLastUpdate & "," & _
                       " status = " & VStatus & "," & _
                       " usercode = " & Vusercode & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    cleargrid grdMatAdjust
    InitGrid
    FillGrid
    initMemvars
    On Error Resume Next
    txtMatCde.SetFocus
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
    If Function_Access(LOGID, "Acess_ADD") = False Then Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            initMemvars
            cmdMatAdjust2.ZOrder 1
            picMatAdjust2.ZOrder 1
        Case vbKeyF2
            AddorEdit = "ADD"
            cmdMatAdjust2.ZOrder 0
            picMatAdjust2.ZOrder 0
            initMemvars
            txtMatCde.Enabled = True
            On Error Resume Next
            txtMatCde.SetFocus
        Case vbKeyF3
            grdMatAdjust.Col = 0
            If grdMatAdjust.Text = "No Entry" Or grdMatAdjust.Text = "TAG NO." Then
                MsgBoxXP "Nothing to Edit!", "Empty Record", XP_OKOnly, msg_Information
                Exit Sub
            End If
            AddorEdit = "EDIT"
            cmdMatAdjust2.ZOrder 0
            picMatAdjust2.ZOrder 0
            initMemvars
            StoreMemVars
        Case vbKeyF4
            cmdDelete_Click
        Case vbKeyF6
            Unload Me
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    cleargrid grdMatAdjust
    rsRefresh
    initMemvars
    InitGrid
    FillGrid
    cmdMatAdjust2.ZOrder 1
    picMatAdjust2.ZOrder 1
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsMatAdjust = New ADODB.Recordset
    rsMatAdjust.Open "Select * from CSMS_MatAdjust order by lastupdate asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    cleargrid grdMatAdjust
    With grdMatAdjust
        .Row = 0
        .FormatString = "Material Code        | Material Description         | Cost           |   Add     |   Minus   | Last Update | User Code | Status     | Particular                                                        "
        .ColWidth(9) = 1
    End With
End Sub

Sub FillGrid()
    Dim kcnt                                           As Integer
    kcnt = 0
    Dim VSTATUSTEXT                                    As String
    If Not rsMatAdjust.EOF And Not rsMatAdjust.BOF Then
        Screen.MousePointer = 11
        rsMatAdjust.MoveFirst
        Do While Not rsMatAdjust.EOF
            kcnt = kcnt + 1
            If Null2String(rsMatAdjust!Status) = "N" Then VSTATUSTEXT = Null2String(rsMatAdjust!Status) Else VSTATUSTEXT = "POSTED"
            grdMatAdjust.AddItem Null2String(rsMatAdjust!MATCDE) & Chr(9) & _
                                 Null2String(rsMatAdjust!MatDsc) & Chr(9) & _
                                 N2Str2Zero(rsMatAdjust!COST) & Chr(9) & _
                                 N2Str2Zero(rsMatAdjust![Add]) & Chr(9) & _
                                 N2Str2Zero(rsMatAdjust!minus) & Chr(9) & _
                                 Null2String(rsMatAdjust!lastupdate) & Chr(9) & _
                                 Null2String(rsMatAdjust!usercode) & Chr(9) & _
                                 VSTATUSTEXT & Chr(9) & _
                                 Null2String(rsMatAdjust!particular) & Chr(9) & _
                                 rsMatAdjust!ID
            DoEvents
            If VSTATUSTEXT = "POSTED" Then
                grdMatAdjust.Row = kcnt + 1
                grdMatAdjust.Col = 7
                grdMatAdjust.CellForeColor = vbWhite
                grdMatAdjust.CellBackColor = vbRed
            End If
            rsMatAdjust.MoveNext
        Loop
        If kcnt <> 0 Then grdMatAdjust.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Sub initMemvars()
    txtMatCde.Text = ""
    txtCost.Text = 0
    txtAdd.Text = 0
    txtMinus.Text = 0
    txtParticular.Text = ""
    cmdSave.Enabled = False
End Sub

Sub StoreMemVars()
    grdMatAdjust.Row = grdMatAdjust.Row
    grdMatAdjust.Col = 9
    Set rsMatAdjust = New ADODB.Recordset
    rsMatAdjust.Open "Select * from CSMS_MatAdjust where id=" & NumericVal(grdMatAdjust.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatAdjust.EOF And Not rsMatAdjust.BOF Then
        labid.Caption = rsMatAdjust!ID
        txtMatCde.Text = Null2String(rsMatAdjust!MATCDE)
        labMatDsc.Caption = Null2String(rsMatAdjust!MatDsc)
        txtCost.Text = N2Str2Zero(rsMatAdjust!COST)
        txtAdd.Text = N2Str2Zero(rsMatAdjust![Add])
        txtMinus.Text = N2Str2Zero(rsMatAdjust!minus)
        txtParticular.Text = Null2String(rsMatAdjust!particular)
        If Null2String(rsMatAdjust!Status) = "P" Then
            MsgSpeechBox "Warning: Adjustments in this Material Code has been Posted!" & vbCrLf & _
                       "         Changes in this Data has been Disabled."
            cmdCancel_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub grdMatAdjust_DblClick()
    grdMatAdjust.Col = 0
    If grdMatAdjust.Text = "No Entry" Or grdMatAdjust.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    cmdMatAdjust2.ZOrder 0
    picMatAdjust2.ZOrder 0
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

Private Sub txtMatCde_LostFocus()
    If txtMatCde.Text = "" Then Exit Sub
    Dim rsMatAdjustDUP                                 As ADODB.Recordset
    Dim rsMatMas                                       As ADODB.Recordset
    'If AddorEdit = "ADD" Then
    '   Set rsMatAdjustDUP = New ADODB.Recordset
    '       rsMatAdjustDUP.Open "Select MatCde from CSMS_MatAdjust where MatCde=" & N2Str2Null(txtMatCde.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
    '   If Not rsMatAdjustDUP.EOF And Not rsMatAdjustDUP.BOF Then
    '      MsgSpeechBox "Error: This Material Code " & txtMatCde.Text & " is already Adjusted."
    '      cmdSave.Enabled = False
    '      On Error Resume Next
    '      txtMatCde.SetFocus
    '      Exit Sub
    '   End If
    'End If
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select onhand,MatCde,MatDsc,Cost,location from CSMS_MatMas where MatCde = " & N2Str2Null(txtMatCde.Text), gconDMIS
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        txtCost.Text = N2Str2Zero(rsMatMas!COST)
        labMatDsc.Caption = Null2String(rsMatMas!MatDsc)
        cmdSave.Enabled = True
        DoEvents
    Else
        MsgSpeechBox "Error: This Material Code " & txtMatCde.Text & " doesn't exist in Materials Master File."
        labMatDsc.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        txtMatCde.SetFocus
    End If
End Sub
