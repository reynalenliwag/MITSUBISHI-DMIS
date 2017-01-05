VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCSMS_BayInformation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCSMS_BayInformation.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListJobs 
      Height          =   2265
      Left            =   90
      TabIndex        =   6
      Top             =   3900
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3995
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
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Flate Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Std Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Technician"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView ListAcc 
      Height          =   2265
      Left            =   90
      TabIndex        =   22
      Top             =   3900
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accessories Description "
         Object.Width           =   8468
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tatal Amount"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView listParts 
      Height          =   2265
      Left            =   90
      TabIndex        =   21
      Top             =   3900
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Parts Description "
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total Amount"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView listMat 
      Height          =   2265
      Left            =   90
      TabIndex        =   23
      Top             =   3900
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Material Description"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Amount"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2190
      TabIndex        =   20
      Top             =   2550
      Width           =   4365
   End
   Begin VB.Label LblSA 
      BackStyle       =   0  'Transparent
      Caption         =   "Sa:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2220
      TabIndex        =   19
      Top             =   2220
      Width           =   5025
   End
   Begin VB.Label LblVehicle 
      BackStyle       =   0  'Transparent
      Caption         =   "VehMode:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2190
      TabIndex        =   18
      Top             =   1590
      Width           =   6795
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2190
      TabIndex        =   17
      Top             =   900
      Width           =   7335
   End
   Begin VB.Label LblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2190
      TabIndex        =   16
      Top             =   1230
      Width           =   7365
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2190
      TabIndex        =   15
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Time/Date:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   13
      Top             =   2550
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Advisor:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   12
      Top             =   2220
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   11
      Top             =   1590
      Width           =   1725
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Plate no:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   10
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Address:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   1230
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   8
      Top             =   900
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   7
      Top             =   300
      Width           =   1575
   End
   Begin VB.Label lbljobs 
      BackStyle       =   0  'Transparent
      Caption         =   "View Jobs"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   510
      TabIndex        =   5
      Top             =   3480
      Width           =   1635
   End
   Begin VB.Label lblMat 
      BackStyle       =   0  'Transparent
      Caption         =   "View Material"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   7140
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblacc 
      BackStyle       =   0  'Transparent
      Caption         =   "View Accesories"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   4710
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblparts 
      BackStyle       =   0  'Transparent
      Caption         =   "View Parts"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   2790
      TabIndex        =   2
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label lblPlate 
      BackStyle       =   0  'Transparent
      Caption         =   "Plate"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   1
      Top             =   1920
      Width           =   3105
   End
   Begin VB.Label lblRO 
      BackStyle       =   0  'Transparent
      Caption         =   "RO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      TabIndex        =   0
      Top             =   300
      Width           =   2325
   End
End
Attribute VB_Name = "FrmCSMS_BayInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theRO As String
Dim thePlateNo As String
Private Sub Form_GotFocus()
    theRO = lblRO
    thePlateNo = lblPlate
End Sub
Private Sub Form_Load()
   Call CenterMe(frmMain, Me, 1)
   theRO = lblRO
   thePlateNo = lblPlate
   lbljobs_Click
End Sub
Private Sub lblacc_Click()
    ListJobs.Visible = False
    listParts.Visible = False
    ListAcc.Visible = True
    listMat.Visible = False
    theRO = lblRO.Caption
    thePlateNo = lblPlate.Caption
    DisplayACC
End Sub
Private Sub lbljobs_Click()
    ListJobs.Visible = True
    listParts.Visible = False
    ListAcc.Visible = False
    listMat.Visible = False
    theRO = lblRO.Caption
    thePlateNo = lblPlate.Caption
    DisplayJobs
End Sub
Private Sub lblMat_Click()
     ListJobs.Visible = False
    listParts.Visible = False
    ListAcc.Visible = False
    listMat.Visible = True
    theRO = lblRO.Caption
    thePlateNo = lblPlate.Caption
    DisplayMaterial
End Sub
Private Sub lblparts_Click()
    ListJobs.Visible = False
    listParts.Visible = True
    ListAcc.Visible = False
    listMat.Visible = False
    theRO = lblRO.Caption
    thePlateNo = lblPlate.Caption
    displayParts
End Sub
Sub DisplayMaterial()
    Dim rsUpload  As New ADODB.Recordset
    Dim item As ListItem
    listParts.Sorted = False: listParts.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='2' AND REP_OR = '" & theRO & "' Order by [LINE_NO] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Do While Not rsUpload.EOF
            Set item = listParts.ListItems.Add(, , Null2String(rsUpload!DetCDE))
            item.SubItems(1) = Null2String(rsUpload!Detdsc)
            item.SubItems(2) = Null2String(rsUpload!detvol)
            item.SubItems(3) = Null2String(rsUpload!DetPrc)
            item.SubItems(4) = Null2String(rsUpload!Det_AMT)
            rsUpload.MoveNext
        Loop
    End If
End Sub
Sub displayParts()
    Dim rsUpload As New ADODB.Recordset
     Dim item As ListItem
    listParts.Sorted = False: listParts.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='2' AND REP_OR = '" & theRO & "' Order by [LINE_NO] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Do While Not rsUpload.EOF
            Set item = listParts.ListItems.Add(, , Null2String(rsUpload!DetCDE))
            item.SubItems(1) = Null2String(rsUpload!Detdsc)
            item.SubItems(2) = Null2String(rsUpload!detvol)
            item.SubItems(3) = Null2String(rsUpload!DetPrc)
            item.SubItems(4) = Null2String(rsUpload!Det_AMT)
            rsUpload.MoveNext
        Loop
    End If

End Sub
Sub DisplayACC()
    Dim rsUpload  As New ADODB.Recordset
    Dim item As ListItem
    ListAcc.Sorted = False: ListAcc.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='4' AND REP_OR = '" & theRO & "' Order by [LINE_NO] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Do While Not rsUpload.EOF
            Set item = ListAcc.ListItems.Add(, , Null2String(rsUpload!DetCDE))
            item.SubItems(1) = Null2String(rsUpload!Detdsc)
            item.SubItems(2) = Null2String(rsUpload!detvol)
            item.SubItems(3) = Null2String(rsUpload!DetPrc)
            item.SubItems(4) = Null2String(rsUpload!Det_AMT)
            rsUpload.MoveNext
        Loop
    End If
End Sub
Sub DisplayJobs()
    Dim rsUpload  As New ADODB.Recordset
    Dim item As ListItem
    ListJobs.Sorted = False: ListJobs.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select DetCDE,Detdsc,FlatRate,Det_Hrs,Technician from CSMS_Ro_Det where LIVIL='1' AND REP_OR = '" & theRO & "' Order by [LINE_NO] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Do While Not rsUpload.EOF
            Set item = ListJobs.ListItems.Add(, , Null2String(rsUpload!DetCDE))
            item.SubItems(1) = Null2String(rsUpload!Detdsc)
            item.SubItems(2) = Null2String(rsUpload!FlatRate)
            item.SubItems(3) = Null2String(rsUpload!Det_Hrs)
            item.SubItems(4) = Null2String(rsUpload!Technician)
            rsUpload.MoveNext
        Loop
    End If
End Sub

