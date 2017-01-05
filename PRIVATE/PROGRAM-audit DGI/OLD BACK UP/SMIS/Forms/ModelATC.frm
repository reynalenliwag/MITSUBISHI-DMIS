VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Files_ModelATC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Models Account Code"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ModelATC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   5940
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
      Height          =   795
      Left            =   5010
      MouseIcon       =   "ModelATC.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "ModelATC.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit Window"
      Top             =   4140
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
      Left            =   4320
      MouseIcon       =   "ModelATC.frx":0D82
      MousePointer    =   99  'Custom
      Picture         =   "ModelATC.frx":0ED4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Save this Record"
      Top             =   4140
      Width           =   705
   End
   Begin VB.ComboBox cboV_Model 
      ForeColor       =   &H00701E2A&
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   150
      Width           =   4305
   End
   Begin VB.Frame Frame3 
      Caption         =   "Purchase && Inventory"
      Height          =   1515
      Left            =   180
      TabIndex        =   2
      Top             =   510
      Width           =   5655
      Begin VB.CommandButton Command11 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   34
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   33
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   32
         Top             =   1050
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   24
         Top             =   1050
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txt_AtcInventoryUnit 
         Height          =   375
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   15
         Tag             =   "Inventory Unit"
         Top             =   1020
         Width           =   2685
      End
      Begin VB.TextBox txt_AtcSalesUnit_CostofSalesFleet 
         Height          =   375
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   14
         Tag             =   "Cost of Sales Fleet"
         Top             =   570
         Width           =   2685
      End
      Begin VB.TextBox txt_AtcSalesUnit_CostofSalesRetail 
         Height          =   375
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "Cost of Sales Retail"
         Top             =   150
         Width           =   2685
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost of Sales Fleet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost of Sales Retail"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Unit"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   1020
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sales"
      Height          =   1935
      Left            =   180
      TabIndex        =   6
      Top             =   1980
      Width           =   5625
      Begin VB.CommandButton Command15 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   38
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   37
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-"
         Height          =   315
         Left            =   5160
         TabIndex        =   35
         Top             =   1500
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   28
         Top             =   1500
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "::"
         Height          =   315
         Left            =   4800
         TabIndex        =   25
         Top             =   210
         Width           =   375
      End
      Begin VB.TextBox txt_AtcSalesUnit_Fleet 
         Height          =   375
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Tag             =   "Sales Unit - Fleet"
         Top             =   1470
         Width           =   2655
      End
      Begin VB.TextBox txt_AtcSalesUnit_Retail 
         Height          =   375
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         Tag             =   "Sales Unit -Retail"
         Top             =   1020
         Width           =   2655
      End
      Begin VB.TextBox txt_AtcSalesDisc_Fleet 
         Height          =   375
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         Tag             =   "Sales Discount - Fleet"
         Top             =   630
         Width           =   2655
      End
      Begin VB.TextBox txt_AtcSalesDisc_Retail 
         Height          =   375
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         Tag             =   "Sales Discount - Retail"
         Top             =   210
         Width           =   2655
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Unit -Retail"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Unit - Fleet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   1995
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Discount - Retail"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Discount - Fleet"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search For Account Code"
      Height          =   5085
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   5835
      Begin XtremeReportControl.ReportControl ListView1 
         Height          =   3915
         Left            =   90
         TabIndex        =   31
         Top             =   1020
         Width           =   5655
         _Version        =   655364
         _ExtentX        =   9975
         _ExtentY        =   6906
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton Command8 
         Caption         =   "X"
         Height          =   285
         Left            =   5280
         TabIndex        =   29
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   90
         MaxLength       =   20
         TabIndex        =   21
         Top             =   570
         Width           =   5145
      End
      Begin VB.Label labAccountDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   30
         Top             =   240
         Width           =   5445
      End
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Model "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -1080
      TabIndex        =   0
      Top             =   150
      Width           =   2055
   End
End
Attribute VB_Name = "frmSMIS_Files_ModelATC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsModel                                                           As ADODB.Recordset
Dim txt                                                               As TextBox
Private Sub cboV_Model_Change()
    rsModel.MoveFirst
    rsModel.Find ("MODEL='" & cboV_Model & "'")
    StoreMemVars
End Sub

Private Sub cboV_Model_Click()
    cboV_Model_Change
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:23
Private Sub cmdSave_Click()
    Dim sql                                                           As String
    '
    sql = "UPDATE ALL_MODEL SET "
    sql = sql & "ATC_SALESDISC_FLEET=" & N2Str2Null((txt_AtcSalesDisc_Fleet)) & " ,"
    sql = sql & " ATC_SALESDISC_RETAIL=" & N2Str2Null((txt_AtcSalesDisc_Retail)) & " ,"
    sql = sql & " ATC_SALES_FLEET=" & N2Str2Null((txt_AtcSalesUnit_Fleet)) & " ,"
    sql = sql & " ATC_SALES_RETAIL=" & N2Str2Null((txt_AtcSalesUnit_Retail)) & " ,"
    sql = sql & " ATC_COSTOFSALES_FLEET=" & N2Str2Null((txt_AtcSalesUnit_CostofSalesFleet)) & " ,"
    sql = sql & " ATC_COSTOFSALES_RETAIL=" & N2Str2Null((txt_AtcSalesUnit_CostofSalesRetail)) & " ,"
    sql = sql & " ATC_INVENTORY=" & N2Str2Null((txt_AtcInventoryUnit))
    sql = sql & " WHERE MODEL=" & N2Str2Null(cboV_Model)
    gconDMIS.Execute sql
    MessagePop RecSaveOk, "Record Updated", "Record Sucessfully Updated", 1000
    rsModel.Requery
    rsModel.Find ("MODEL='" & cboV_Model & "'")
    StoreMemVars
End Sub

Function GetATCCode(xxx)
    Dim rsATCCODE                                                     As ADODB.Recordset
    GetATCCode = ""
    If xxx <> "" Then
        Set rsATCCODE = gconDMIS.Execute("SELECT ACCTCODE FROM AMIS_ChartAccount WHERE DESCRIPTION=" & N2Str2Null(xxx))
        If Not (rsATCCODE.EOF Or rsATCCODE.BOF) Then
            GetATCCode = Null2String(rsATCCODE!ACCTCODE)
        End If
    End If
End Function

Function SetATCCode(xxx)
    On Error GoTo ErrorCode
    Dim rsATCCODE                                                     As ADODB.Recordset
    SetATCCode = ""
    If xxx <> "" Then
        Set rsATCCODE = gconDMIS.Execute("SELECT description FROM AMIS_ChartAccount WHERE ACCTCODE=" & N2Str2Null(xxx))
        If Not (rsATCCODE.EOF Or rsATCCODE.BOF) Then
            SetATCCode = Null2String(rsATCCODE!Description)
        End If
    End If
    Exit Function
ErrorCode:
    ShowVBError
End Function

Private Sub Command1_Click()
    Set txt = txt_AtcSalesUnit_CostofSalesRetail
    BringToFront
End Sub

Private Sub Command10_Click()
    txt_AtcSalesUnit_CostofSalesFleet = ""
End Sub

Private Sub Command11_Click()
    txt_AtcSalesUnit_CostofSalesRetail = ""
End Sub

Private Sub Command12_Click()
    txt_AtcSalesUnit_Fleet = ""
End Sub

Private Sub Command13_Click()
    txt_AtcSalesUnit_Retail = ""
End Sub

Private Sub Command14_Click()
    txt_AtcSalesDisc_Fleet = ""
End Sub

Private Sub Command15_Click()
    txt_AtcSalesDisc_Retail = ""
End Sub

Private Sub Command2_Click()
    Set txt = txt_AtcSalesUnit_CostofSalesFleet
    BringToFront
End Sub

Private Sub Command3_Click()
    Set txt = txt_AtcInventoryUnit
    BringToFront
End Sub

Private Sub Command4_Click()
    Set txt = txt_AtcSalesDisc_Retail
    BringToFront
End Sub

Private Sub Command5_Click()
    Set txt = txt_AtcSalesDisc_Fleet
    BringToFront
End Sub

Private Sub Command6_Click()
    Set txt = txt_AtcSalesUnit_Retail
    BringToFront
End Sub

Private Sub Command7_Click()
    Set txt = txt_AtcSalesUnit_Fleet
    BringToFront
End Sub

Private Sub Command8_Click()
    SendToBack
End Sub

Private Sub Command9_Click()
    txt_AtcInventoryUnit = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
Sub BringToFront()
    Frame1.Visible = True: Frame1.ZOrder 0
    labAccountDetail = txt.Tag & "-" & cboV_Model
    Text1.SetFocus
End Sub
Sub SendToBack()
    Frame1.Visible = False: Frame1.ZOrder 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Combo_Loadval cboV_Model, gconDMIS.Execute("SELECT  DISTINCT MODEL  from ALL_MODEL  order by MODEL ASC")

    rsRefresh
    initMemvars
    StoreMemVars

    FillGrid
    If cboV_Model.ListCount > 0 Then
        cboV_Model.ListIndex = 0
    End If
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txt_AtcInventoryUnit = ""
    txt_AtcSalesDisc_Fleet = ""
    txt_AtcSalesDisc_Retail = ""
    txt_AtcSalesUnit_CostofSalesRetail = ""
    txt_AtcSalesUnit_Fleet = ""
    txt_AtcSalesUnit_Retail = ""
End Sub

Sub rsRefresh()
    Set rsModel = New ADODB.Recordset
    rsModel.Open "SELECT  * from ALL_MODEL  order by MODEL DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub
'

Sub StoreMemVars()
    If Not rsModel.EOF And Not rsModel.BOF Then
        txt_AtcInventoryUnit = Null2String(rsModel("ATC_INVENTORY"))

        txt_AtcSalesUnit_CostofSalesFleet = Null2String(rsModel("ATC_COSTOFSALES_Fleet"))
        txt_AtcSalesUnit_CostofSalesRetail = Null2String(rsModel("ATC_COSTOFSALES_RETAIL"))



        txt_AtcSalesDisc_Fleet = Null2String(rsModel("ATC_SALESDISC_FLEET"))
        txt_AtcSalesDisc_Retail = Null2String(rsModel("ATC_SALESDISC_RETAIL"))

        txt_AtcSalesUnit_Fleet = Null2String(rsModel("ATC_SALES_FLEET"))
        txt_AtcSalesUnit_Retail = Null2String(rsModel("ATC_SALES_RETAIL"))
    End If
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorCode
    If KeyCode = 13 Then
        If ListView1.Rows.Count > 0 Then
            txt.Text = ListView1.SelectedRows.Row(0).Record(0).Value
            txt.SetFocus
            Set txt = Nothing
            SendToBack
        End If

    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub ListView1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    ListView1_KeyDown 13, 1
End Sub

Private Sub Text1_Change()
    ListView1.FilterText = Text1
    ListView1.Populate
End Sub
Sub FillGrid()
    On Error GoTo ErrorCode
    Dim temprs                                                        As ADODB.Recordset
    Dim DepCode                                                       As String
    Dim rsdep                                                         As ADODB.Recordset
    ReportControlAddColumnHeader ListView1, "Account Code, Description"
    ResizeColumnHeader ListView1, "40,55"
    ReportControlPaintManager ListView1

    '  Set rsdep = gconDMIS.Execute("Select DeptCode from AMIS_Department where DEPTNAME ='SALES DEPARTMENT'")

    'If Not rsdep.EOF Or Not rsdep.BOF Then
    '   DepCode = Null2String(rsdep!DeptCode)
    'End If

    '    If DepCode = "" Then
    Set temprs = gconDMIS.Execute("Select AcctCode,Description from AMIS_ChartAccount order by 2 ASC")
    '    Else
    '   Set TempRs = gconDMIS.Execute("Select AcctCode,Description from AMIS_ChartAccount where departmentcode='" & DepCode & " ' order by 2 ASC")
    '   End If

    If Not temprs.EOF Or Not temprs.BOF Then
        flex_FillReportView temprs, ListView1, False
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        ListView1.SetFocus
    End If
End Sub

Private Sub txt_AtcInventoryUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command3_Click
End Sub

Private Sub txt_AtcSalesDisc_Fleet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command5_Click
End Sub

Private Sub txt_AtcSalesDisc_Retail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command4_Click
End Sub

Private Sub txt_AtcSalesUnit_CostofSalesFleet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command2_Click
End Sub

Private Sub txt_AtcSalesUnit_CostofSalesRetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command1_Click
End Sub

Private Sub txt_AtcSalesUnit_Fleet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command7_Click
End Sub

Private Sub txt_AtcSalesUnit_Retail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: Command6_Click
End Sub
