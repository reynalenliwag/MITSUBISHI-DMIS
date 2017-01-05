VERSION 5.00
Begin VB.Form frmCSMSAddVehAppt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   3585
   ClientLeft      =   450
   ClientTop       =   645
   ClientWidth     =   7740
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmAddVehAppt.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   6060
      ScaleHeight     =   900
      ScaleWidth      =   1605
      TabIndex        =   24
      Top             =   2640
      Width           =   1605
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
         Left            =   840
         MouseIcon       =   "FrmAddVehAppt.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "FrmAddVehAppt.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         Left            =   120
         MouseIcon       =   "FrmAddVehAppt.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "FrmAddVehAppt.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save Vehicle"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   7605
      Begin VB.TextBox txtTin_Number 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5670
         MaxLength       =   15
         TabIndex        =   27
         Top             =   1320
         Width           =   1845
      End
      Begin VB.TextBox txtDel_Date 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5670
         MaxLength       =   18
         TabIndex        =   10
         Top             =   1710
         Width           =   1845
      End
      Begin VB.TextBox txtD_Sold 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5670
         MaxLength       =   18
         TabIndex        =   9
         Top             =   540
         Width           =   1845
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
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
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Text            =   "cboModel"
         Top             =   1320
         Width           =   1845
      End
      Begin VB.ComboBox cboModel 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
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
         Height          =   315
         Left            =   1830
         TabIndex        =   7
         Text            =   "cboModel"
         Top             =   930
         Width           =   1845
      End
      Begin VB.TextBox txtWar_Cert 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5670
         MaxLength       =   15
         TabIndex        =   6
         Top             =   930
         Width           =   1845
      End
      Begin VB.TextBox txtSerial 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5670
         MaxLength       =   18
         TabIndex        =   5
         Top             =   150
         Width           =   1845
      End
      Begin VB.TextBox txtEngine 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   18
         TabIndex        =   4
         Top             =   1680
         Width           =   1845
      End
      Begin VB.TextBox txtVCond_No 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   3
         Top             =   540
         Width           =   1845
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   2
         Top             =   150
         Width           =   1845
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Delivered"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   1740
         Width           =   1665
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   1380
         Width           =   1665
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Certificate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   990
         Width           =   1965
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   210
         Width           =   2085
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1740
         Width           =   1635
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   1020
         Width           =   1995
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Conduction Sticker"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   600
         Width           =   2085
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.TextBox txtPlate_No 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   300
      TabIndex        =   23
      Top             =   630
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   630
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Plate Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   21
      Top             =   150
      Width           =   2025
   End
End
Attribute VB_Name = "frmCSMSAddVehAppt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer, rsCusVeh, rsS_Model, rsColor           As ADODB.Recordset
Attribute rsCusVeh.VB_VarUserMemId = 1073938432
Attribute rsS_Model.VB_VarUserMemId = 1073938432
Attribute rsColor.VB_VarUserMemId = 1073938432
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938436

Function SetColor(CCC As String)
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function SetColorDesc(CCC As String)
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select * from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColorDesc = Null2String(rsColor!color_desc)
    Else
        SetColorDesc = ""
    End If
End Function

Sub initMemvars()
    txtProdNo.Text = ""
    txtVCond_No.Text = ""
    cboModel.Text = ""
    cboCOLOR.Text = ""
    txtENGINE.Text = ""
    txtSerial.Text = ""
    txtD_Sold.Text = ""
    txtWar_Cert.Text = ""
    txtTin_Number.Text = ""
    txtDel_Date.Text = ""
    FillCbo
End Sub

Sub StoreMemVars()
    On Error GoTo ErrorCode
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        labid.Caption = rsCusVeh!ID
        txtProdNo.Text = Null2String(rsCusVeh!ProdNo)
        txtVCond_No.Text = Null2String(rsCusVeh!VCOND_NO)
        cboModel.Text = Null2String(rsCusVeh!MODEL)
        cboCOLOR.Text = SetColorDesc(Null2String(rsCusVeh!ClrCde))
        txtENGINE.Text = Null2String(rsCusVeh!Engine)
        txtSerial.Text = Null2String(rsCusVeh!SERIAL)
        txtD_Sold.Text = Null2String(rsCusVeh!D_SOLD)
        txtWar_Cert.Text = Null2String(rsCusVeh!War_Cert)
        txtTin_Number.Text = Null2String(rsCusVeh!TIN_Number)
        txtDel_Date.Text = Null2String(rsCusVeh!DEL_DATE)
        AddorEdit = "EDIT"
    Else
        initMemvars
        AddorEdit = "ADD"
    End If

    Exit Sub

ErrorCode:

    ShowVBError

End Sub

Sub FillCbo()
    Set rsS_Model = New ADODB.Recordset
    rsS_Model.Open "Select model from CSMS_S_Model", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboModel.Clear
        Do While Not rsS_Model.EOF
            cboModel.AddItem Null2String(rsS_Model!MODEL)
            rsS_Model.MoveNext
        Loop
    End If
    Set rsColor = New ADODB.Recordset
    rsColor.Open "Select COLOR_DESC from ALL_Color", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsColor.EOF And Not rsColor.BOF Then
        rsColor.MoveFirst
        cboCOLOR.Clear
        Do While Not rsColor.EOF
            cboCOLOR.AddItem Null2String(rsColor!color_desc)
            rsColor.MoveNext
        Loop
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()


    On Error GoTo ErrorCode
    Dim vtxtCusCde, VTXTNiym, VTXTPlateNo              As String

    Dim VtxtProdNo, VtxtVCond_No, VcboModel            As String
    Dim Vcbocolor, VtxtEngine, VtxtSerial              As String
    Dim VtxtD_Sold, VtxtWar_Cert, VtxtTin_Number       As String
    Dim VtxtDel_Date                                   As String

    If txtPlate_No = "" Then
        MsgBox "Plate Number please..."
        Exit Sub
    End If
    VTXTPlateNo = N2Str2Null(txtPlate_No.Text)
    VtxtProdNo = N2Str2Null(txtProdNo.Text)
    VtxtVCond_No = N2Str2Null(txtVCond_No.Text)
    VcboModel = N2Str2Null(cboModel.Text)
    Vcbocolor = N2Str2Null(SetColor(cboCOLOR.Text))
    VtxtEngine = N2Str2Null(txtENGINE.Text)
    VtxtSerial = N2Str2Null(txtSerial.Text)
    VtxtD_Sold = N2Str2Null(Format(txtD_Sold.Text, "short date"))
    VtxtWar_Cert = N2Str2Null(txtWar_Cert.Text)
    VtxtTin_Number = N2Str2Null(txtTin_Number.Text)
    VtxtDel_Date = N2Str2Null(Format(txtDel_Date.Text, "Short date"))

    If AddorEdit = "ADD" Then
        If IsNull(txtProdNo.Text) = False Then
            Dim rsCusVehDup                            As ADODB.Recordset
            Set rsCusVehDup = New ADODB.Recordset
            rsCusVehDup.Open "select prodno from CSMS_CusVeh where prodno = '" & txtProdNo.Text & "'", gconDMIS
            If Not rsCusVehDup.EOF And Not rsCusVeh.BOF Then
                MsgSpeechBox "Product Number Already Exist"
                Exit Sub
            End If
        End If
        gconDMIS.Execute "insert into CSMS_CusVeh " & _
                         "(cuscde,niym,plate_no,prodno,vcond_no,model,clrcde,engine,serial,tin_number,d_sold,war_cert,del_date)" & _
                       " values (" & vtxtCusCde & ", " & VTXTNiym & ", " & VTXTPlateNo & ", " & VtxtProdNo & ", " & VtxtVCond_No & ", " & VcboModel & ", " & Vcbocolor & ", " & VtxtEngine & ", " & VtxtSerial & ", " & VtxtTin_Number & ", " & VtxtD_Sold & ", " & VtxtWar_Cert & ", " & VtxtDel_Date & ")"
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select id,cuscde from ALL_Customer where cuscde = " & vtxtCusCde, gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            gconDMIS.Execute "update ALL_Customer set plate_no = " & VTXTPlateNo & " where id = " & rsCustomer!ID
        End If
    Else
        gconDMIS.Execute "update CSMS_CusVeh set" & _
                       " prodno = " & VtxtProdNo & ", " & _
                       " vcond_no = " & VtxtVCond_No & ", " & _
                       " model = " & VcboModel & ", " & _
                       " clrcde = " & Vcbocolor & ", " & _
                       " engine = " & VtxtEngine & ", " & _
                       " serial = " & VtxtSerial & ", " & _
                       " tin_number = " & VtxtTin_Number & ", " & _
                       " d_sold = " & VtxtD_Sold & ", " & _
                       " war_cert = " & VtxtWar_Cert & ", " & _
                       " del_date = " & VtxtDel_Date & _
                       " where id = " & labid.Caption
    End If
    LogAudit "A", "CUSTOMER VEHICLES", " PLATE#:" & txtPlate_No
    Unload Me
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    initMemvars
    Set rsCusVeh = New ADODB.Recordset
    rsCusVeh.Open "select * from CSMS_CusVeh where plate_no = 'saga'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    txtPlate_No.Text = ""
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSROCusveh = Nothing
End Sub

Private Sub txtD_Sold_LostFocus()
    If txtD_Sold.Text <> "" Then txtD_Sold.Text = Format(txtD_Sold.Text, "Short Date")
End Sub

Private Sub txtDel_Date_LostFocus()
    If txtDel_Date.Text <> "" Then txtDel_Date.Text = Format(txtDel_Date.Text, "Short Date")
End Sub

Private Sub txtProdNo_Change()

    Dim rsSMIS_MRRINV                                  As ADODB.Recordset
    Dim rsPurchAgree                                   As ADODB.Recordset
    Set rsSMIS_MRRINV = New ADODB.Recordset
    Set rsSMIS_MRRINV = gconDMIS.Execute("Select * from SMIS_MRRINV Where ProdNo = " & N2Str2Null(txtProdNo.Text))
    If Not rsSMIS_MRRINV.EOF And Not rsSMIS_MRRINV.BOF Then
        txtVCond_No.Text = Null2String(rsSMIS_MRRINV!ignkey)
        cboModel.Text = Null2String(rsSMIS_MRRINV!MODEL)
        cboCOLOR.Text = Null2String(rsSMIS_MRRINV!Color)
        txtENGINE.Text = Null2String(rsSMIS_MRRINV!EngineNo)
        txtSerial.Text = Null2String(rsSMIS_MRRINV!SerialNo)
        txtD_Sold.Text = Null2String(rsSMIS_MRRINV!DateReleased)
        txtDel_Date.Text = Null2String(rsSMIS_MRRINV!DateReceived)
    End If

End Sub

Private Sub txtVCond_No_Change()
    Dim rsSMIS_MRRINV                                  As ADODB.Recordset
    Dim rsPurchAgree                                   As ADODB.Recordset
    Set rsSMIS_MRRINV = New ADODB.Recordset
    Set rsSMIS_MRRINV = gconDMIS.Execute("Select * from SMIS_MRRINV Where IGNKey = " & N2Str2Null(txtVCond_No.Text))
    If Not rsSMIS_MRRINV.EOF And Not rsSMIS_MRRINV.BOF Then
        txtProdNo.Text = Null2String(rsSMIS_MRRINV!ProdNo)
        cboModel.Text = Null2String(rsSMIS_MRRINV!MODEL)
        cboCOLOR.Text = Null2String(rsSMIS_MRRINV!Color)
        txtENGINE.Text = Null2String(rsSMIS_MRRINV!EngineNo)
        txtSerial.Text = Null2String(rsSMIS_MRRINV!SerialNo)
        txtD_Sold.Text = Null2String(rsSMIS_MRRINV!DateReleased)
        txtDel_Date.Text = Null2String(rsSMIS_MRRINV!DateReceived)

    End If
End Sub

