VERSION 5.00
Begin VB.Form frmSMIS_Files_CustomerVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Vehicles"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerVehicles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6015
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Height          =   6210
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   30
      TabIndex        =   41
      Top             =   0
      Width           =   30
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   960
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   5955
      TabIndex        =   30
      Top             =   6210
      Width           =   6015
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   3075
         ScaleHeight     =   900
         ScaleWidth      =   5295
         TabIndex        =   34
         Top             =   0
         Width           =   5295
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   30
            MouseIcon       =   "CustomerVehicles.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   755
            MouseIcon       =   "CustomerVehicles.frx":0D7B
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":0ECD
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2595
            MouseIcon       =   "CustomerVehicles.frx":1225
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":1377
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   30
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   2910
            MouseIcon       =   "CustomerVehicles.frx":168A
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":17DC
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   30
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   1500
            MouseIcon       =   "CustomerVehicles.frx":1B07
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":1C59
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   1515
            MouseIcon       =   "CustomerVehicles.frx":1FBF
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":2111
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   30
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   4320
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   31
         Top             =   45
         Width           =   2580
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   30
            MouseIcon       =   "CustomerVehicles.frx":246D
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":25BF
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   765
            MouseIcon       =   "CustomerVehicles.frx":290F
            MousePointer    =   99  'Custom
            Picture         =   "CustomerVehicles.frx":2A61
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.Label labNoOfVehciles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "There Are 3 Units Registerd For This Customer"
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6210
      Left            =   30
      ScaleHeight     =   6210
      ScaleWidth      =   5805
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      Begin VB.ComboBox cboCLRCode 
         Height          =   330
         Left            =   2115
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   4875
         Width           =   3675
      End
      Begin VB.TextBox txtVCOND_NO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   26
         Top             =   5655
         Width           =   3645
      End
      Begin VB.TextBox txtVin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   25
         Top             =   5235
         Width           =   3645
      End
      Begin VB.TextBox txtTIN_NUMBER 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   22
         Top             =   4455
         Width           =   3645
      End
      Begin VB.TextBox txtKMReading 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   15
         Top             =   1830
         Width           =   3645
      End
      Begin VB.TextBox txtProductNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   14
         Top             =   2265
         Width           =   3645
      End
      Begin VB.TextBox txtSerialNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   13
         Top             =   2700
         Width           =   3645
      End
      Begin VB.TextBox txtDateSold 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   12
         Top             =   3135
         Width           =   3645
      End
      Begin VB.TextBox txtWarrantyCertificateNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   11
         Top             =   3570
         Width           =   3645
      End
      Begin VB.TextBox txtSellingDealer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   10
         Top             =   4005
         Width           =   3645
      End
      Begin VB.TextBox txtEngine 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   7
         Top             =   1380
         Width           =   3645
      End
      Begin VB.TextBox txtModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   5
         Top             =   960
         Width           =   3645
      End
      Begin VB.TextBox txtyear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3945
         TabIndex        =   4
         Top             =   510
         Width           =   1785
      End
      Begin VB.TextBox txtPlateNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   2
         Top             =   510
         Width           =   1755
      End
      Begin VB.TextBox txtCustomerName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2115
         TabIndex        =   1
         Top             =   60
         Width           =   3615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Conduction Number"
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
         Index           =   14
         Left            =   405
         TabIndex        =   29
         Top             =   5715
         Width           =   1680
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "VIN Number"
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
         Index           =   13
         Left            =   1065
         TabIndex        =   28
         Top             =   5325
         Width           =   1005
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Color"
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
         Index           =   12
         Left            =   1620
         TabIndex        =   24
         Top             =   4965
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tin Number"
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
         Index           =   11
         Left            =   1095
         TabIndex        =   23
         Top             =   4575
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Last KM Reading"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   495
         TabIndex        =   21
         Top             =   1890
         Width           =   1650
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Product Number"
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
         Left            =   720
         TabIndex        =   20
         Top             =   2355
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Serial Number"
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
         Left            =   870
         TabIndex        =   19
         Top             =   2805
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Date Sold"
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
         Left            =   1260
         TabIndex        =   18
         Top             =   3210
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Warranty Certificate No"
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
         Left            =   105
         TabIndex        =   17
         Top             =   3660
         Width           =   1995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selling Dealer"
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
         Index           =   10
         Left            =   900
         TabIndex        =   16
         Top             =   4095
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "CustomerName"
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
         Index           =   0
         Left            =   735
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Engine"
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
         Left            =   1485
         TabIndex        =   8
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Model"
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
         Left            =   1575
         TabIndex        =   6
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Plate No/ Year"
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
         Left            =   900
         TabIndex        =   3
         Top             =   600
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmSMIS_Files_CustomerVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsVehicles                         As ADODB.Recordset

Private Sub cmdAdd_Click()
    picAdds.Visible = True: picSaves.Visible = False: Picture3.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True: picSaves.Visible = False: Picture3.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    picAdds.Visible = False: picSaves.Visible = True: Picture3.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Not RsVehicles.EOF Then
        RsVehicles.MoveNext
    End If

    If RsVehicles.EOF And RsVehicles.RecordCount > 0 Then
        RsVehicles.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()

    If Not RsVehicles.BOF Then
        RsVehicles.MovePrevious
    End If

    If RsVehicles.BOF And RsVehicles.RecordCount > 0 Then
        RsVehicles.MoveFirst
        ShowFirstRecordMsg
    End If

    StoreMemvars
End Sub

Private Sub cmdSave_Click()
    With RsVehicles
        If AddorEdit = "ADD" Then
            .AddNew
        End If
    End With
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    picSaves.Visible = False
    picAdds.Visible = True
End Sub

Sub ShowVehicles(xCuscde As String)
    Set RsVehicles = New ADODB.Recordset

    Call RsVehicles.Open("Select * from CSMS_CUSVEH where CUSCDE=" & N2Str2Null(xCuscde), gconDMIS, adOpenDynamic, adLockOptimistic)

    If RsVehicles.RecordCount > 1 Then
        labNoOfVehciles.Caption = " There Are " & RsVehicles.RecordCount & " Registered Vehicles For " & txtCustomerName
        picAdds.Visible = True
        picSaves.Visible = False
        'labNoOfVehciles.Visible = True
        '  cmdNext.Visible = True
        ' cmdPrevious.Visible = True
    Else
        labNoOfVehciles.Caption = ""
        picSaves.Visible = True
        picAdds.Visible = False
        'labNoOfVehciles.Visible = False
        'cmdNext.Visible = False
        'cmdPrevious.Visible = False
    End If
    StoreMemvars

End Sub

Sub StoreMemvars()
'ID
'CUSCDE
'NIYM
'VIN
'PLATE_NO
'VCOND_NO
'CLRCDE
'YER
'MAKE
'MODEL
'ENGINE
'KMREADING
'PRODNO
'SERIAL
'TIN_NUMBER
'D_SOLD
'WAR_CERT
'DEL_DATE
'SELLING_DEALER

    If Not (RsVehicles.EOF Or RsVehicles.BOF) Then
        txtCustomerName = Null2String(RsVehicles("NIYM"))
        txtDateSold = Null2String(RsVehicles("D_SOLD"))
        txtEngine = Null2String(RsVehicles("ENGINE"))
        txtKMReading = Null2String(RsVehicles("KMreading"))
        txtModel = Null2String(RsVehicles("MODEL"))
        txtPlateNo = Null2String(RsVehicles("PLATE_NO"))
        txtProductNumber = Null2String(RsVehicles("PRODNO"))
        txtSellingDealer = Null2String(RsVehicles("SELLING_DEALER"))
        txtWarrantyCertificateNo = Null2String(RsVehicles("WAR_CERT"))
        txtYear = Null2String(RsVehicles("YER"))
        txtVin = Null2String(RsVehicles("VIN"))
        txtVCOND_NO = Null2String(RsVehicles("VCOND_NO"))
        cboCLRCode = Null2String(RsVehicles("CLRCDE"))
        txtTIN_NUMBER = Null2String(RsVehicles("TIN_NUMBER"))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

