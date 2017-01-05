VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTranTypeImportingSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importing Set-up"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12180
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTranTypeImportingSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12180
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3330
      ScaleHeight     =   1095
      ScaleWidth      =   1485
      TabIndex        =   31
      Top             =   7500
      Width           =   1485
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
         Left            =   720
         MouseIcon       =   "frmTranTypeImportingSetup.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmTranTypeImportingSetup.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cancel"
         Top             =   150
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
         Left            =   30
         MouseIcon       =   "frmTranTypeImportingSetup.frx":1512
         MousePointer    =   99  'Custom
         Picture         =   "frmTranTypeImportingSetup.frx":1664
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Save Entry"
         Top             =   150
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   4380
      ScaleHeight     =   2745
      ScaleWidth      =   7755
      TabIndex        =   13
      Top             =   5850
      Width           =   7755
      Begin VB.TextBox txtARAccount 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2280
         Width           =   1785
      End
      Begin VB.TextBox txtARDescription 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_INV_DESCRIPT 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1170
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_COGS_DESCRIPT 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   780
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_DESCRIPT 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   390
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_INV_CODE 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1170
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_COGS_CODE 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   780
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_CODE 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   18
         Tag             =   "D"
         Top             =   390
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_DIS_DESCRIPT 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1920
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_OTC_DESCRIPT 
         Height          =   375
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_DIS_CODE 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_OTC_CODE 
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label lblAccountCode 
         Height          =   225
         Left            =   90
         TabIndex        =   37
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1320
         TabIndex        =   36
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   3555
         TabIndex        =   30
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1695
         TabIndex        =   29
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inventory A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   465
         TabIndex        =   28
         Top             =   1230
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Sales A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sales A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   765
         TabIndex        =   26
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   825
         TabIndex        =   25
         Top             =   1980
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Output Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   645
         TabIndex        =   24
         Top             =   1620
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   30
      ScaleHeight     =   2745
      ScaleWidth      =   4245
      TabIndex        =   1
      Top             =   5850
      Width           =   4245
      Begin VB.TextBox txtTranType4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1080
         MaxLength       =   35
         TabIndex        =   12
         Top             =   2220
         Width           =   2085
      End
      Begin VB.TextBox txtTranType3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1080
         MaxLength       =   35
         TabIndex        =   11
         Top             =   1830
         Width           =   2085
      End
      Begin VB.TextBox txtTranType2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1080
         MaxLength       =   35
         TabIndex        =   10
         Top             =   1440
         Width           =   2085
      End
      Begin VB.TextBox txtTranType1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1080
         MaxLength       =   35
         TabIndex        =   9
         Top             =   1020
         Width           =   2085
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   4
         Top             =   480
         Width           =   4035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   150
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TranType4:"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   2310
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TranType3:"
         Height          =   225
         Left            =   90
         TabIndex        =   7
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TranType2:"
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   1500
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TranType1:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lvChart 
      Height          =   5745
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   10134
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TranType1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TranType2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TranType3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TranType4"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTranTypeImportingSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vTranType1                                     As String
    Dim vTranType2                                     As String
    Dim vTranType3                                     As String
    Dim vTranType4                                     As String
    vTranType1 = N2Str2Null(txtTranType1.Text)
    vTranType2 = N2Str2Null(txtTranType2.Text)
    vTranType3 = N2Str2Null(txtTranType3.Text)
    vTranType4 = N2Str2Null(txtTranType4.Text)
    If lblAccountCode.Caption = "" Then
        MsgBox "Please select account", vbInformation, "Account Code"
        Exit Sub
    Else
        If MsgBox("Update Account Code " & lblAccountCode.Caption & Chr(13) & "Description " & lvChart.SelectedItem.SubItems(1), vbQuestion + vbYesNo, "Update?") = vbYes Then
            gconDMIS.Execute "UPDATE AMIS_CHARTACCOUNT SET TRANTYPE1 = " & vTranType1 & " , TRANTYPE2 = " & vTranType2 & ", TRANTYPE3 = " & vTranType3 & " , TRANTYPE4 = " & vTranType4 & " WHERE ACCTCODE = '" & lblAccountCode.Caption & "'"
            Call ChartofAccountsList(txtSearch)
            lblAccountCode.Caption = ""
            MsgBox "Selected Account Code updated", vbInformation, "Record Updated"
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    ChartofAccountsList
End Sub

Private Sub ChartofAccountsList(Optional XXX As String)
    Dim xList                                          As ListItem
    lvChart.ListItems.Clear
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    If Option1.Value = True Then
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where DESCRIPTION LIKE '%" & XXX & "%'", gconDMIS, adOpenKeyset
    Else
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where ACCTCODE LIKE '" & XXX & "%'", gconDMIS, adOpenKeyset
    End If
    If Not rsChart.EOF And Not rsChart.BOF Then
        Do While Not rsChart.EOF
            Set xList = lvChart.ListItems.Add(, , Null2String(rsChart!ACCTCODE))
            xList.SubItems(1) = Null2String(rsChart!DESCRIPTION)
            xList.SubItems(2) = Null2String(rsChart!TRANTYPE1)
            xList.SubItems(3) = Null2String(rsChart!Trantype2)
            xList.SubItems(4) = Null2String(rsChart!Trantype3)
            xList.SubItems(5) = Null2String(rsChart!Trantype4)
            rsChart.MoveNext
        Loop
    End If
    Set rsChart = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Label1_Click(Index As Integer)
    Dim xList                                          As ListItem
    lvChart.ListItems.Clear
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where TRANTYPE1 = '" & txtTranType1.Text & "'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        Do While Not rsChart.EOF
            Set xList = lvChart.ListItems.Add(, , Null2String(rsChart!ACCTCODE))
            xList.SubItems(1) = Null2String(rsChart!DESCRIPTION)
            xList.SubItems(2) = Null2String(rsChart!TRANTYPE1)
            xList.SubItems(3) = Null2String(rsChart!Trantype2)
            xList.SubItems(4) = Null2String(rsChart!Trantype3)
            xList.SubItems(5) = Null2String(rsChart!Trantype4)
            rsChart.MoveNext
        Loop
    End If
    Set rsChart = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Label2_Click()
    Dim xList                                          As ListItem
    lvChart.ListItems.Clear
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where TRANTYPE2 = '" & txtTranType2.Text & "'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        Do While Not rsChart.EOF
            Set xList = lvChart.ListItems.Add(, , Null2String(rsChart!ACCTCODE))
            xList.SubItems(1) = Null2String(rsChart!DESCRIPTION)
            xList.SubItems(2) = Null2String(rsChart!TRANTYPE1)
            xList.SubItems(3) = Null2String(rsChart!Trantype2)
            xList.SubItems(4) = Null2String(rsChart!Trantype3)
            xList.SubItems(5) = Null2String(rsChart!Trantype4)
            rsChart.MoveNext
        Loop
    End If
    Set rsChart = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Label3_Click()
    Dim xList                                          As ListItem
    lvChart.ListItems.Clear
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where TRANTYPE3 = '" & txtTranType3.Text & "'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        Do While Not rsChart.EOF
            Set xList = lvChart.ListItems.Add(, , Null2String(rsChart!ACCTCODE))
            xList.SubItems(1) = Null2String(rsChart!DESCRIPTION)
            xList.SubItems(2) = Null2String(rsChart!TRANTYPE1)
            xList.SubItems(3) = Null2String(rsChart!Trantype2)
            xList.SubItems(4) = Null2String(rsChart!Trantype3)
            xList.SubItems(5) = Null2String(rsChart!Trantype4)
            rsChart.MoveNext
        Loop
    End If
    Set rsChart = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Label4_Click()
    Dim xList                                          As ListItem
    lvChart.ListItems.Clear
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT where TRANTYPE4 = '" & txtTranType4.Text & "'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        Do While Not rsChart.EOF
            Set xList = lvChart.ListItems.Add(, , Null2String(rsChart!ACCTCODE))
            xList.SubItems(1) = Null2String(rsChart!DESCRIPTION)
            xList.SubItems(2) = Null2String(rsChart!TRANTYPE1)
            xList.SubItems(3) = Null2String(rsChart!Trantype2)
            xList.SubItems(4) = Null2String(rsChart!Trantype3)
            xList.SubItems(5) = Null2String(rsChart!Trantype4)
            rsChart.MoveNext
        Loop
    End If
    Set rsChart = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub lvChart_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblAccountCode.Caption = lvChart.SelectedItem.Text
    txtTranType1.Text = lvChart.SelectedItem.SubItems(2)
    txtTranType2.Text = lvChart.SelectedItem.SubItems(3)
    txtTranType3.Text = lvChart.SelectedItem.SubItems(4)
    txtTranType4.Text = lvChart.SelectedItem.SubItems(5)

    Call ReturnSales(txtTranType2.Text, txtTranType1.Text)
    Call ReturnCostofSales(txtTranType2.Text, txtTranType1.Text)
    Call ReturnInventory(txtTranType2.Text)
    Call ReturnOutputTax("OUTPUT TAX")
    Call ReturnDiscount(txtTranType2.Text)
    Call ReturnAR(txtTranType2.Text)
End Sub

Private Sub Option1_Click()
    txtSearch.Text = ""
End Sub

Private Sub Option2_Click()
    txtSearch.Text = ""
End Sub

Private Sub txtSearch_Change()

    Call ChartofAccountsList(txtSearch)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ReturnSales(xTRANTYPE2 As String, Optional xTRANTYPE1 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    If txtTranType1 = "" Then
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'SALES'", gconDMIS, adOpenKeyset
    Else
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = '" & xTRANTYPE1 & "' AND TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'SALES'", gconDMIS, adOpenKeyset
    End If
    If Not rsChart.EOF And Not rsChart.BOF Then
        txt_SALES_ACT_CODE.Text = Null2String(rsChart!ACCTCODE)
        txt_SALES_ACT_DESCRIPT.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txt_SALES_ACT_CODE.Text = ""
        txt_SALES_ACT_DESCRIPT.Text = ""
    End If
End Sub

Private Sub ReturnCostofSales(xTRANTYPE2 As String, Optional xTRANTYPE1 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    If txtTranType1 = "" Then
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'COST OF SALES'", gconDMIS, adOpenKeyset
    Else
        rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = '" & xTRANTYPE1 & "' AND TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'COST OF SALES'", gconDMIS, adOpenKeyset
    End If
    If Not rsChart.EOF And Not rsChart.BOF Then
        txt_SALES_ACT_COGS_CODE.Text = Null2String(rsChart!ACCTCODE)
        txt_SALES_ACT_COGS_DESCRIPT.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txt_SALES_ACT_COGS_CODE.Text = ""
        txt_SALES_ACT_COGS_DESCRIPT.Text = ""
    End If
End Sub

Private Sub ReturnInventory(xTRANTYPE2 As String, Optional xTRANTYPE1 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'INVENTORY'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        txt_SALES_ACT_INV_CODE.Text = Null2String(rsChart!ACCTCODE)
        txt_SALES_ACT_INV_DESCRIPT.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txt_SALES_ACT_INV_CODE.Text = ""
        txt_SALES_ACT_INV_DESCRIPT.Text = ""
    End If
End Sub

Private Sub ReturnOutputTax(xTRANTYPE1 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = '" & xTRANTYPE1 & "'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        txt_SALES_ACT_OTC_CODE.Text = Null2String(rsChart!ACCTCODE)
        txt_SALES_ACT_OTC_DESCRIPT.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txt_SALES_ACT_OTC_CODE.Text = ""
        txt_SALES_ACT_OTC_DESCRIPT.Text = ""
    End If
End Sub

Private Sub ReturnDiscount(xTRANTYPE2 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE2 = '" & xTRANTYPE2 & "' AND TRANTYPE3 = 'DISCOUNT'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        txt_SALES_ACT_DIS_CODE.Text = Null2String(rsChart!ACCTCODE)
        txt_SALES_ACT_DIS_DESCRIPT.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txt_SALES_ACT_DIS_CODE.Text = ""
        txt_SALES_ACT_DIS_DESCRIPT.Text = ""
    End If
End Sub

Private Sub ReturnAR(xTRANTYPE1 As String)
    Dim rsChart                                        As ADODB.Recordset
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT ACCTCODE,DESCRIPTION,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4 FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = '" & xTRANTYPE1 & "' AND TRANTYPE2='AR'", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        txtARAccount.Text = Null2String(rsChart!ACCTCODE)
        txtARDescription.Text = Null2String(rsChart!DESCRIPTION)
    Else
        txtARAccount.Text = ""
        txtARDescription.Text = ""
    End If
End Sub

Private Sub txtTranType1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTranType2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTranType3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTranType4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
