VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "wizMacForm.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAllTOOLS 
   BorderStyle     =   0  'None
   Caption         =   "DMIS 2.0 System Administrator Tool"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   10451
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Administration"
      TabPicture(0)   =   "AllTOOLS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture7"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PMIS"
      TabPicture(1)   =   "AllTOOLS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CSMS"
      TabPicture(2)   =   "AllTOOLS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "SMIS"
      TabPicture(3)   =   "AllTOOLS.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "AMIS"
      TabPicture(4)   =   "AllTOOLS.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "CMIS"
      TabPicture(5)   =   "AllTOOLS.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture2"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "HRMS"
      TabPicture(6)   =   "AllTOOLS.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Picture1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00000000&
         Height          =   4665
         Left            =   -74910
         ScaleHeight     =   4605
         ScaleWidth      =   8505
         TabIndex        =   7
         Top             =   120
         Width           =   8565
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   -74940
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   6
         Top             =   90
         Width           =   8565
         Begin wizButton.cmd cmdReUpdateTranDetails 
            Height          =   435
            Left            =   4860
            TabIndex        =   16
            Top             =   1740
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Re-Update Parts Tran Details"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":00C4
         End
         Begin wizButton.cmd cmdSetGenuineNonGenuine 
            Height          =   435
            Left            =   90
            TabIndex        =   9
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Set Genuine / Non-Genuine"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":00E0
         End
         Begin wizButton.cmd cmdUploadMaterials 
            Height          =   435
            Left            =   90
            TabIndex        =   10
            Top             =   1740
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Upload Material"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":00FC
         End
         Begin wizButton.cmd cmdUploadParts 
            Height          =   435
            Left            =   90
            TabIndex        =   11
            Top             =   2250
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Upload Parts"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0118
         End
         Begin wizButton.cmd cmdProcessPartsCost 
            Height          =   435
            Left            =   4860
            TabIndex        =   12
            Top             =   180
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Process Parts Cost"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0134
         End
         Begin wizButton.cmd cmdMatCost 
            Height          =   435
            Left            =   4860
            TabIndex        =   13
            Top             =   720
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Process Mat Cost"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0150
         End
         Begin wizButton.cmd cmdAccCost 
            Height          =   435
            Left            =   4860
            TabIndex        =   14
            Top             =   1230
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Process Acc Cost"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":016C
         End
         Begin wizButton.cmd cmdSetStockMasHARINONHARI 
            Height          =   435
            Left            =   90
            TabIndex        =   15
            Top             =   720
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Set StockMas Hari / Non-Hari"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0188
         End
         Begin wizButton.cmd cmdReUpdateTranType 
            Height          =   435
            Left            =   4860
            TabIndex        =   20
            Top             =   2250
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Re-Update Parts Tran Type"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":01A4
         End
         Begin wizButton.cmd cmdReUpdateTranCost 
            Height          =   435
            Left            =   4860
            TabIndex        =   21
            Top             =   2790
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Re-Update Parts Tran Cost"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":01C0
         End
         Begin wizButton.cmd cmdStorePartsBegBal 
            Height          =   435
            Left            =   90
            TabIndex        =   27
            Top             =   1230
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Store Beg Balance"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":01DC
         End
         Begin wizButton.cmd cmd1 
            Height          =   435
            Left            =   120
            TabIndex        =   28
            Top             =   2760
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Set Item No"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":01F8
         End
         Begin wizButton.cmd cmdUploadAccessories 
            Height          =   435
            Left            =   120
            TabIndex        =   29
            Top             =   3240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Upload Material"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0214
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   -74910
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   5
         Top             =   120
         Width           =   8565
         Begin wizButton.cmd cmdSetCustCodeInCusVeh 
            Height          =   435
            Left            =   180
            TabIndex        =   18
            Top             =   390
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Set CustCode in CusVeh"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0230
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   -74910
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   4
         Top             =   120
         Width           =   8565
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   -74910
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   3
         Top             =   120
         Width           =   8565
         Begin wizButton.cmd cmdRefreshVendorCode 
            Height          =   435
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Refresh Vendor Code"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":024C
         End
         Begin wizButton.cmd cmdRefreshCustCode 
            Height          =   435
            Left            =   120
            TabIndex        =   19
            Top             =   810
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   767
            TX              =   "Refresh Customer Code"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0268
         End
         Begin wizButton.cmd cmdCheckLostEntries 
            Height          =   435
            Left            =   5370
            TabIndex        =   22
            Top             =   180
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            TX              =   "Check Lost Entries"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":0284
         End
         Begin wizButton.cmd cmdCheckUnBalEntries 
            Height          =   435
            Left            =   5370
            TabIndex        =   23
            Top             =   720
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            TX              =   "Check Un-Balance Entries"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":02A0
         End
         Begin wizButton.cmd cmdInvalidCustCode 
            Height          =   435
            Left            =   5370
            TabIndex        =   24
            Top             =   1260
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            TX              =   "Invalid Customer Code"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":02BC
         End
         Begin wizButton.cmd cmdInvalidJournalEntry 
            Height          =   435
            Left            =   5370
            TabIndex        =   25
            Top             =   1800
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            TX              =   "Invalid Journal Entry"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":02D8
         End
         Begin wizButton.cmd cmdInvalidJournalStatus 
            Height          =   435
            Left            =   5370
            TabIndex        =   26
            Top             =   2310
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            TX              =   "Invalid Journal Status"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AllTOOLS.frx":02F4
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   -74910
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   2
         Top             =   120
         Width           =   8565
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   90
         ScaleHeight     =   4665
         ScaleWidth      =   8565
         TabIndex        =   1
         Top             =   120
         Width           =   8565
      End
   End
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   556
      MacCaption      =   "DMIS TOOLS"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
End
Attribute VB_Name = "frmAllTOOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function GetCustomerCode(lastname As String) As String
    Dim temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = Left(Trim(lastname), 1)
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerCode = Left(lastname, 1) & "00001"
    End If
End Function

Function GetCustomerZCode(lastname As String) As String
    Dim temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = "Z"
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetCustomerZCode = lAlpha & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerZCode = lAlpha & "00001"
    End If
End Function

Function GetVendorCode(lastname As String) As String
    Dim temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = Left(Trim(lastname), 1)
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_VENCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetVendorCode = Left(lastname, 1) & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetVendorCode = Left(lastname, 1) & "00001"
    End If
End Function

Function GetVendorZCode(lastname As String) As String
    Dim temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = "Z"
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_VENCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetVendorZCode = lAlpha & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetVendorZCode = lAlpha & "00001"
    End If
End Function

Function SetHariOrNonHari(XXX As String) As String
    Dim rsSTOCKMAS                                     As ADODB.Recordset
    Set rsSTOCKMAS = New ADODB.Recordset
    Set rsSTOCKMAS = gconDMIS.Execute("Select NON_HARI from PMIS_STOCKMAS WHERE STOCKNO = '" & XXX & "'")
    If Not rsSTOCKMAS.EOF And Not rsSTOCKMAS.BOF Then
        SetHariOrNonHari = Null2String(rsSTOCKMAS!NON_HARI)
    Else
        SetHariOrNonHari = "D"
    End If
End Function

Function SetTYPE(XXX As String) As String
    Dim rsSTOCKMAS                                     As ADODB.Recordset
    Set rsSTOCKMAS = New ADODB.Recordset
    Set rsSTOCKMAS = gconDMIS.Execute("Select TYPE from PMIS_STOCKMAS WHERE STOCKNO = '" & XXX & "'")
    If Not rsSTOCKMAS.EOF And Not rsSTOCKMAS.BOF Then
        SetTYPE = Null2String(rsSTOCKMAS!Type)
    Else
        SetTYPE = "D"
    End If
End Function

Private Sub cmd1_Click()

    Dim rsTran                                         As ADODB.Recordset
    Dim RSHDR                                          As ADODB.Recordset

    Set RSHDR = gconDMIS.Execute("SELECT TRANTYPE, TRANNO, Type  FROM PMIS_ORD_HIST WHERE TRANTYPE IN('CSH','ADB','RIV','CHG') ORDER BY ID DESC")
    DoEvents
    While Not RSHDR.EOF
        DoEvents
        frmMain.Caption = RSHDR!TRANNO
        Set rsTran = gconDMIS.Execute("SELECT  ID " & _
                                    " FROM PMIS_DAYTRAN WHERE " & _
                                    " TRANTYPE =" & N2Str2Null(RSHDR!TRANTYPE) & _
                                    " AND TRANNO =" & N2Str2Null(RSHDR!TRANNO) & _
                                    " AND TYPE=" & N2Str2Null(RSHDR!Type) & _
                                    " ORDER BY ID ASC ")
        i = 0

        While Not rsTran.EOF
            i = i + 1
            gconDMIS.Execute ("UPDATE PMIS_DAYTRAN SET ITEMNO='" & Format(i, "0000") & "' WHERE ID=" & rsTran!ID)
            DoEvents
            frmMain.Caption = Format(i, "0000") & "---" & N2Str2Null(RSHDR!Type) & "---------------" & N2Str2Null(RSHDR!TRANTYPE)
            rsTran.MoveNext
        Wend

        RSHDR.MoveNext
    Wend


    MsgBox "ok"





End Sub

Private Sub cmdCheckLostEntries_Click()
    Dim rsJournal_Det                                  As ADODB.Recordset
    Dim rsJournal_HD                                   As ADODB.Recordset

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det order by JNO ASC")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Do While Not rsJournal_Det.EOF
            Set rsJournal_HD = New ADODB.Recordset
            Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_journal_HD where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO))
            If rsJournal_HD.EOF And rsJournal_HD.BOF Then
                MsgBox "Invalid Data" & vbCrLf & _
                       Null2String(rsJournal_Det!jtype) & "-" & Null2String(rsJournal_Det!VOUCHERNO)
                gconDMIS.Execute ("Delete from AMIS_Journal_Det where ID = " & rsJournal_Det!ID)
            End If
            Me.Caption = Null2String(rsJournal_Det!jtype) & "-" & Null2String(rsJournal_Det!VOUCHERNO): DoEvents
            wizMacApp1.MacCaption = Me.Caption: DoEvents
            rsJournal_Det.MoveNext
        Loop
    End If
    MsgBox "ok"
End Sub

Private Sub cmdCheckUnBalEntries_Click()
    Dim rsJournal_HD                                   As ADODB.Recordset
    Dim rsJournal_Det                                  As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where jtype <> 'OPB' AND STATUS = 'P' Order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_HD.EOF
            Set rsJournal_Det = New ADODB.Recordset
            Set rsJournal_Det = gconDMIS.Execute("Select SUM(DEBIT) AS TotalDebit,SUM(CREDIT) AS TotalCredit from AMIS_Journal_Det Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " And VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " AND STATUS = 'P'")
            If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
                If Round(N2Str2Zero(rsJournal_Det!TotalDebit), 2) <> Round(N2Str2Zero(rsJournal_Det!Totalcredit), 2) Then
                    gconDMIS.Execute "update AMIS_Journal_HD set " & _
                                   " Debit = " & N2Str2Zero(rsJournal_Det!TotalDebit) & "," & _
                                   " Credit = " & N2Str2Zero(rsJournal_Det!Totalcredit) & "," & _
                                   " Status = 'N'" & _
                                   " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)
                    gconDMIS.Execute "update AMIS_Journal_Det set " & _
                                   " Status = 'N'" & _
                                   " Where JType = " & N2Str2Null(rsJournal_HD!jtype) & " AND VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO)

                    'MsgBox Null2String(rsJournal_HD!jTYPE) & "-" & Null2String(rsJournal_HD!vOUCHERNO) & " is not Balance."
                Else
                    'gconDMIS.Execute "update AMIS_Journal_HD set " & _
                     "Debit = " & N2Str2Zero(rsJOURNAL_DET!TotalDebit) & "," & _
                     "Credit = " & N2Str2Zero(rsJOURNAL_DET!Totalcredit) & _
                     " Where JType = " & N2Str2Null(rsJOURNAL_HD!JType) & " AND VoucherNo = " & N2Str2Null(rsJOURNAL_HD!VoucherNo)
                End If
            Else
                MsgBox Null2String(rsJournal_HD!jtype) & "-" & Null2String(rsJournal_HD!VOUCHERNO) & " has no Detail."
            End If
            Me.Caption = Null2String(rsJournal_HD!VOUCHERNO)
            rsJournal_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "Completed!"
End Sub

Private Sub cmdInvalidCustCode_Click()
    Dim rsJournal_HD                                   As ADODB.Recordset
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Order by id asc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_HD.EOF
            Set rsCustomer = New ADODB.Recordset
            Set rsCustomer = gconDMIS.Execute("select * from ALL_CustMaster_Amis where CustCode <> '999999' AND CustCode = " & N2Str2Null(rsJournal_HD!CustomerCode))
            If rsCustomer.EOF And rsCustomer.BOF Then
                MsgBox Null2String(rsJournal_HD!CustomerCode) & " is Invalid!"
            End If
            rsJournal_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "Completed!"
End Sub

Private Sub cmdInvalidJournalEntry_Click()
    Dim rsJournal_Det                                  As ADODB.Recordset
    Dim rsChartAccounts                                As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            gconDMIS.Execute ("Update AMIS_Journal_Det set JNo = " & N2Str2Null(rsJournal_Det!JNo) & ", Jdate = " & N2Date2Null(rsJournal_Det!JDate) & ", status = " & N2Str2Null(rsJournal_Det!Status) & " where VoucherNo = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and Jtype = " & N2Str2Null(rsJournal_Det!jtype))
            Me.Caption = rsJournal_Det!JNo: DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "ok"
    Exit Sub

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'P' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'P'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'P' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'N' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'N'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'N' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Un-Posted Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_HD where status = 'C' order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccounts = New ADODB.Recordset
            Set rsChartAccounts = gconDMIS.Execute("Select * from AMIS_Journal_det where Jtype = " & N2Str2Null(rsJournal_Det!jtype) & " and voucherno = " & N2Str2Null(rsJournal_Det!VOUCHERNO) & " and status <> 'C'")
            If Not rsChartAccounts.EOF And Not rsChartAccounts.BOF Then
                gconDMIS.Execute ("update AMIS_Journal_Det set status = 'C' where id = " & rsChartAccounts!ID)
                'MsgBox "Invalid Cancelled Transaction" & vbCrLf & _
                 Null2String(rsChartAccounts!acct_code) & " " & Null2String(rsChartAccounts!acct_Name) & vbCrLf & _
                 Null2String(rsJOURNAL_DET!Jtype) & "-" & Null2String(rsJOURNAL_DET!voucherno)
            End If
            Me.Caption = Null2String(rsJournal_Det!JNo): DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    MsgBox "Completed!"

End Sub

Private Sub cmdInvalidJournalStatus_Click()
    Dim rsJournal_Det                                  As ADODB.Recordset
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("Select ID,Voucherno,jtype,acct_code,acct_name,Jno,status from AMIS_Journal_Det Order by id asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        rsJournal_Det.MoveFirst
        MsgBox "Poon"
        Screen.MousePointer = 11
        Do While Not rsJournal_Det.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select jno,status from AMIS_Journal_HD Where status <> " & N2Str2Null(rsJournal_Det!Status) & " and JNo = " & N2Str2Null(rsJournal_Det!JNo))
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                If Null2String(rsChartAccount!Status) = "P" Then
                    gconDMIS.Execute ("update AMIS_Journal_Det SET STATUS = 'P' WHERE JNO = " & N2Str2Null(rsJournal_Det!JNo))
                Else
                    MsgBox "HEADER STATUS = (" & Null2String(rsChartAccount!Status) & ")" & vbCrLf & _
                           "DETAIL STATUS = (" & Null2String(rsJournal_Det!Status) & ")" & vbCrLf & _
                           Null2String(rsJournal_Det!jtype) & "-" & _
                           Null2String(rsJournal_Det!VOUCHERNO) & vbCrLf & _
                           Null2String(rsJournal_Det!acct_Name)
                End If
            End If
            Me.Caption = "[" & rsJournal_Det!ID & "] " & Null2String(rsJournal_Det!VOUCHERNO)
            DoEvents
            rsJournal_Det.MoveNext
        Loop
        Screen.MousePointer = 0
        MsgBox "Tapos"
    End If
End Sub

Private Sub cmdMatCost_Click()
    Dim rsRO_DET                                       As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim RSDAYTRAN                                      As ADODB.Recordset
    Dim RSREPOR                                        As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim VDATE_REL                                      As String
    Dim i                                              As Integer
    Dim IValue                                         As Double
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '3' Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        i = 0
        Do While Not rsRO_DET.EOF
            Set RSORD_HD = New ADODB.Recordset
            Set RSORD_HD = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'M' AND trantype = 'RIV' and RONO = '" & Null2String(rsRO_DET!REP_OR) & "'")
            If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                RSORD_HD.MoveFirst
                Do While Not RSORD_HD.EOF
                    Set RSDAYTRAN = New ADODB.Recordset
                    Set RSDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'M' AND trantype = 'RIV' and tranno = '" & RSORD_HD!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRO_DET!detcde) & " order by trandate desc")
                    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
                        If N2Str2Zero(RSDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & RSDAYTRAN!TRANUCOST & _
                                           " Where id = " & rsRO_DET!ID
                            Me.Caption = "Processing: " & Null2String(rsRO_DET!REP_OR) & " with Detail Amount: " & N2Str2Zero(rsRO_DET!detamt) & " Cost = " & N2Str2Zero(RSDAYTRAN!TRANUCOST)
                        End If
                    End If
                    RSORD_HD.MoveNext
                Loop
            End If
            i = i + 1
            IValue = (i / rsRO_DET.RecordCount) * 100
            Me.Caption = Int(IValue) & "% Completed": DoEvents
            rsRO_DET.MoveNext
        Loop
    End If
End Sub

Private Sub cmdProcessPartsCost_Click()
    Dim rsRO_DET                                       As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim RSDAYTRAN                                      As ADODB.Recordset
    Dim RSREPOR                                        As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim VDATE_REL                                      As String
    Dim i                                              As Integer
    Dim IValue                                         As Double
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '2' Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        i = 0
        Do While Not rsRO_DET.EOF
            Set RSORD_HD = New ADODB.Recordset
            Set RSORD_HD = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'P' AND trantype = 'RIV' and RONO = '" & Null2String(rsRO_DET!REP_OR) & "'")
            If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                RSORD_HD.MoveFirst
                Do While Not RSORD_HD.EOF
                    Set RSDAYTRAN = New ADODB.Recordset
                    Set RSDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'P' AND trantype = 'RIV' and tranno = '" & RSORD_HD!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRO_DET!detcde) & " order by trandate desc")
                    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
                        If N2Str2Zero(RSDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & RSDAYTRAN!TRANUCOST & _
                                           " Where id = " & rsRO_DET!ID
                            Me.Caption = "Processing: " & Null2String(rsRO_DET!REP_OR) & " with Detail Amount: " & N2Str2Zero(rsRO_DET!detamt) & " Cost = " & N2Str2Zero(RSDAYTRAN!TRANUCOST)
                        End If
                    End If
                    RSORD_HD.MoveNext
                Loop
            End If
            i = i + 1
            IValue = (i / rsRO_DET.RecordCount) * 100
            Me.Caption = Int(IValue) & "% Completed": DoEvents
            rsRO_DET.MoveNext
        Loop
    End If
End Sub

Private Sub cmdReUpdateTranCost_Click()
    Dim vTotalTranCost                                 As Double
    Dim RSORD_HD, RSTDAYTRAN                           As ADODB.Recordset
    Dim i, vOrdHDRecNo                                 As Long
    Set RSORD_HD = New ADODB.Recordset
    Dim vTotalQty                                      As Long

    RSORD_HD.Open "select id,trantype,tranno,status from PMIS_vw_ISS_HISTORY where [TYPE] = 'P' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        i = 0
        Me.Caption = "Computing Total Quantity of Request and Issuances..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_AlldayTran where [TYPE] = 'P' AND trantype = " & N2Str2Null(RSORD_HD!TRANTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!TRANQTY))
                    RSTDAYTRAN.MoveNext
                Loop
                If Null2String(RSORD_HD!Status) <> "C" Then
                    gconDMIS.Execute "update PMIS_Ord_HIST set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                    gconDMIS.Execute "update PMIS_Ord_Hd set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                End If
            End If
            DoEvents
            RSORD_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If


    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,trantype,tranno,status from PMIS_vw_ISS_HISTORY where [TYPE] = 'M' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        i = 0
        Me.Caption = "Computing Total Quantity of Request and Issuances..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_AlldayTran where [TYPE] = 'M' AND trantype = " & N2Str2Null(RSORD_HD!TRANTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!TRANQTY))
                    RSTDAYTRAN.MoveNext
                Loop
                If Null2String(RSORD_HD!Status) <> "C" Then
                    gconDMIS.Execute "update PMIS_Ord_HIST set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                    gconDMIS.Execute "update PMIS_Ord_Hd set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                End If
            End If
            i = i + 1
            DoEvents
            RSORD_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If


    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,trantype,tranno,status from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        i = 0
        Me.Caption = "Computing Total Quantity of Request and Issuances..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            DoEvents
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_AlldayTran where [TYPE] = 'A' AND trantype = " & N2Str2Null(RSORD_HD!TRANTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!TRANQTY))
                    RSTDAYTRAN.MoveNext
                Loop
                If Null2String(RSORD_HD!Status) <> "C" Then
                    gconDMIS.Execute "update PMIS_Ord_HIST set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                    gconDMIS.Execute "update PMIS_Ord_Hd set NETCOST = " & vTotalTranCost & ", TotalQty = " & vTotalQty & " where id = " & vOrdHDRecNo
                End If
            End If
            i = i + 1
            DoEvents
            RSORD_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If

    'RE UPDATE TRAN COST
    Dim rsTranDet                                      As ADODB.Recordset
    Dim rsTranDet2                                     As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim vMAC                                           As Double
    Screen.MousePointer = 11
    Set rsTranDet = New ADODB.Recordset
    Set rsTranDet = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'P' and in_out = 'O' order by TRANDATE ASC, TRANTYPE DESC, TRANNO ASC")
    If Not rsTranDet.EOF And Not rsTranDet.BOF Then
        rsTranDet.MoveFirst
        Do While Not rsTranDet.EOF
            Set rsTranDet2 = New ADODB.Recordset
            Set rsTranDet2 = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'P' and TRANTYPE = 'RR' and TRANDATE <= " & N2Date2Null(rsTranDet!trandate) & " AND STOCK_ORD = " & N2Str2Null(rsTranDet!STOCK_ORD) & " order by trandate desc,TRANNO DESC")
            If Not rsTranDet2.EOF And Not rsTranDet2.BOF Then
                vMAC = N2Str2Zero(rsTranDet2!MAC)
            Else
                Set RSPARTMAS = New ADODB.Recordset
                Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_PartMas where PARTNO = " & N2Str2Null(rsTranDet!STOCK_ORD))
                If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                    vMAC = N2Str2Zero(RSPARTMAS!MAC)
                Else
                    vMAC = 0
                End If
            End If
            gconDMIS.Execute ("Update PMIS_TDaytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            gconDMIS.Execute ("Update PMIS_Daytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            rsTranDet.MoveNext
        Loop
    End If

    Set rsTranDet = New ADODB.Recordset
    Set rsTranDet = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'M' and in_out = 'O' order by TRANTYPE DESC, TRANNO ASC")
    If Not rsTranDet.EOF And Not rsTranDet.BOF Then
        rsTranDet.MoveFirst
        Do While Not rsTranDet.EOF
            Set rsTranDet2 = New ADODB.Recordset
            Set rsTranDet2 = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'M' and TRANTYPE = 'RR' and TRANDATE <= " & N2Date2Null(rsTranDet!trandate) & " AND STOCK_ORD = " & N2Str2Null(rsTranDet!STOCK_ORD) & " order by trandate desc,TRANNO DESC")
            If Not rsTranDet2.EOF And Not rsTranDet2.BOF Then
                vMAC = N2Str2Zero(rsTranDet2!MAC)
            Else
                Set RSPARTMAS = New ADODB.Recordset
                Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_StockMas where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTranDet!STOCK_ORD))
                If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                    vMAC = N2Str2Zero(RSPARTMAS!MAC)
                Else
                    vMAC = 0
                End If
            End If
            gconDMIS.Execute ("Update PMIS_TDaytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            gconDMIS.Execute ("Update PMIS_Daytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            rsTranDet.MoveNext
        Loop
    End If

    Set rsTranDet = New ADODB.Recordset
    Set rsTranDet = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'A' and in_out = 'O' order by TRANTYPE DESC, TRANNO ASC")
    If Not rsTranDet.EOF And Not rsTranDet.BOF Then
        rsTranDet.MoveFirst
        Do While Not rsTranDet.EOF
            Set rsTranDet2 = New ADODB.Recordset
            Set rsTranDet2 = gconDMIS.Execute("Select * from PMIS_AllDayTran where TYPE = 'A' and TRANTYPE = 'RR' and TRANDATE <= " & N2Date2Null(rsTranDet!trandate) & " AND STOCK_ORD = " & N2Str2Null(rsTranDet!STOCK_ORD) & " order by trandate desc,TRANNO DESC")
            If Not rsTranDet2.EOF And Not rsTranDet2.BOF Then
                vMAC = N2Str2Zero(rsTranDet2!MAC)
            Else
                Set RSPARTMAS = New ADODB.Recordset
                Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_StockMas where TYPE = 'A' AND STOCKNO = " & N2Str2Null(rsTranDet!STOCK_ORD))
                If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                    vMAC = N2Str2Zero(RSPARTMAS!MAC)
                Else
                    vMAC = 0
                End If
            End If
            gconDMIS.Execute ("Update PMIS_TDaytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            gconDMIS.Execute ("Update PMIS_Daytran set tranucost = " & vMAC & " where id = " & rsTranDet!ID)
            rsTranDet.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "tapos"

End Sub

Private Sub cmdReUpdateTranDetails_Click()
    Dim vTotalQty                                      As Long
    Dim vOrdHDRecNo                                    As Long
    Dim RSPO_HD                                        As ADODB.Recordset
    Dim RSTDAYTRAN                                     As ADODB.Recordset

    Dim IS_TAT_TOS                                     As String
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select id,pono,status,PODATE from PMIS_PO_HD where [TYPE] = 'P' order by pono asc", gconDMIS
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        RSPO_HD.MoveFirst:
        MsgSpeech "Computing Total Quantity of Purchases...": Me.Caption = "Computing Total Quantity of Purchases..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not RSPO_HD.EOF
            vOrdHDRecNo = RSPO_HD!ID
            If Null2String(RSPO_HD!Status) = "" Then
                IS_TAT_TOS = "N"
            Else
                IS_TAT_TOS = Null2String(RSPO_HD!Status)
            End If
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    gconDMIS.Execute "Update PMIS_TdayTran SET TRANDATE = " & N2Date2Null(RSPO_HD!PODATE) & ", STATUS = '" & IS_TAT_TOS & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
            End If
            gconDMIS.Execute ("Update PMIS_PO_HD Set STATUS = '" & IS_TAT_TOS & "' where id = " & RSPO_HD!ID)
            DoEvents
            RSPO_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Dim rsRR_HD                                        As ADODB.Recordset
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select id,rrno,status,RRDATE from PMIS_RR_HD where [TYPE] = 'P' order by rrno asc", gconDMIS
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        MsgSpeech "Computing Total Quantity of Receipts..."
        Me.Caption = "Computing Total Quantity of Receipts..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not rsRR_HD.EOF
            vOrdHDRecNo = rsRR_HD!ID
            If Null2String(RSPO_HD!Status) = "" Then
                IS_TAT_TOS = "N"
            Else
                IS_TAT_TOS = Null2String(rsRR_HD!Status)
            End If
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,status,itemno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    gconDMIS.Execute "Update PMIS_TdayTran SET TRANDATE = " & N2Date2Null(rsRR_HD!RRDATE) & ", STATUS = '" & IS_TAT_TOS & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
            End If
            gconDMIS.Execute ("Update PMIS_RR_HD Set STATUS = '" & IS_TAT_TOS & "' where id = " & rsRR_HD!ID)
            rsRR_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Dim vTotalTranCost                                 As Double
    Dim RSORD_HD                                       As ADODB.Recordset
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "select id,trantype,tranno,status,TRANDATE from PMIS_Ord_HD where [TYPE] = 'P' order by trantype,tranno asc", gconDMIS
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        MsgSpeech "Computing Total Quantity of Request and Issuances..."
        Me.Caption = "Computing Total Quantity of Request and Issuances..."
        Screen.MousePointer = 11
        DoEvents
        Do While Not RSORD_HD.EOF
            vOrdHDRecNo = RSORD_HD!ID
            If Null2String(RSPO_HD!Status) = "" Then
                IS_TAT_TOS = "N"
            Else
                IS_TAT_TOS = Null2String(RSORD_HD!Status)
            End If
            Set RSTDAYTRAN = New ADODB.Recordset
            RSTDAYTRAN.Open "select id,trantype,tranno,tranqty,tranucost,status,itemno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = " & N2Str2Null(RSORD_HD!TRANTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
            If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                RSTDAYTRAN.MoveFirst
                vTotalQty = 0: vTotalTranCost = 0
                Do While Not RSTDAYTRAN.EOF
                    vTotalQty = vTotalQty + N2Str2Zero(RSTDAYTRAN!TRANQTY)
                    vTotalTranCost = vTotalTranCost + (N2Str2Zero(RSTDAYTRAN!TRANUCOST) * N2Str2Zero(RSTDAYTRAN!TRANQTY))
                    gconDMIS.Execute "Update PMIS_TdayTran SET TRANDATE = " & N2Date2Null(RSORD_HD!trandate) & ", STATUS = '" & IS_TAT_TOS & "' where ID = " & RSTDAYTRAN!ID
                    RSTDAYTRAN.MoveNext
                Loop
            End If
            gconDMIS.Execute ("Update PMIS_ORD_HD Set STATUS = '" & IS_TAT_TOS & "' where id = " & RSORD_HD!ID)
            RSORD_HD.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    MsgBox "oki"
End Sub

Private Sub cmdAccCost_Click()
    Dim rsRO_DET                                       As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim RSDAYTRAN                                      As ADODB.Recordset
    Dim RSREPOR                                        As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim VDATE_REL                                      As String
    Dim i                                              As Integer
    Dim IValue                                         As Double
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '4' Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        i = 0
        Do While Not rsRO_DET.EOF
            Set RSORD_HD = New ADODB.Recordset
            Set RSORD_HD = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' AND trantype = 'RIV' and RONO = '" & Null2String(rsRO_DET!REP_OR) & "'")
            If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                RSORD_HD.MoveFirst
                Do While Not RSORD_HD.EOF
                    Set RSDAYTRAN = New ADODB.Recordset
                    Set RSDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'A' AND trantype = 'RIV' and tranno = '" & RSORD_HD!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRO_DET!detcde) & " order by trandate desc")
                    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
                        If N2Str2Zero(RSDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & RSDAYTRAN!TRANUCOST & _
                                           " Where id = " & rsRO_DET!ID
                            Me.Caption = "Processing: " & Null2String(rsRO_DET!REP_OR) & " with Detail Amount: " & N2Str2Zero(rsRO_DET!detamt) & " Cost = " & N2Str2Zero(RSDAYTRAN!TRANUCOST)
                        End If
                    End If
                    RSORD_HD.MoveNext
                Loop
            End If
            i = i + 1
            IValue = (i / rsRO_DET.RecordCount) * 100
            Me.Caption = Int(IValue) & "% Completed": DoEvents
            rsRO_DET.MoveNext
        Loop
    End If
End Sub

Private Sub cmdRefreshCustCode_Click()
    Dim matibayako                                     As ADODB.Recordset
    Set matibayako = New ADODB.Recordset

    Dim rsCustomer                                     As ADODB.Recordset
    Dim k                                              As Integer
    Dim NewCtlCde                                      As String

    Dim kawnter                                        As Integer
    Dim new_customer_code                              As String
    Set matibayako = gconDMIS.Execute("Select * from ALL_Customer_Table order by AcctName asc")
    If Not matibayako.EOF And Not matibayako.BOF Then
        matibayako.MoveFirst: kawnter = 0
        Do While Not matibayako.EOF
            kawnter = kawnter + 1
            Me.Caption = "record number: " & kawnter
            If IsNumeric(Left(Null2String(matibayako!AcctName), 1)) = True Then
                new_customer_code = GetCustomerZCode(Null2String(matibayako!AcctName))
            Else
                new_customer_code = GetCustomerCode(Null2String(matibayako!AcctName))
            End If
            gconDMIS.Execute ("Update All_Customer_Table set cuscde = '" & new_customer_code & "' where ID =" & matibayako!ID)
            Screen.MousePointer = 11
            gconDMIS.Execute "delete from ALL_CusCtl"
            For k = 65 To 90
                Set rsCustomer = New ADODB.Recordset
                rsCustomer.Open "select CusCde from ALL_Customer_Table where left(CusCde,1) = '" & Chr(k) & "' order by CusCde desc", gconDMIS
                If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                    NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!CUSCDE, 2, 5)) + 1, "00000")
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
                Else
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "',' Customer control character for " & Chr(k) & " -')"
                End If
            Next
            Screen.MousePointer = 0
            matibayako.MoveNext
            DoEvents
        Loop
    End If
End Sub

Private Sub cmdRefreshVendorCode_Click()
    Dim matibayako                                     As ADODB.Recordset
    Set matibayako = New ADODB.Recordset

    Dim rsCustomer                                     As ADODB.Recordset
    Dim k                                              As Integer
    Dim NewCtlCde                                      As String

    Dim kawnter                                        As Integer
    Dim new_customer_code                              As String
    Set matibayako = gconDMIS.Execute("Select * from ALL_Vendor order by nameofvendor asc")
    If Not matibayako.EOF And Not matibayako.BOF Then
        matibayako.MoveFirst: kawnter = 0
        Do While Not matibayako.EOF
            kawnter = kawnter + 1
            Me.Caption = "record number: " & kawnter
            If IsNumeric(Left(Null2String(matibayako!lastname), 1)) = True Then
                new_customer_code = GetVendorZCode(Null2String(matibayako!nameofvendor))
            Else
                new_customer_code = GetVendorCode(Null2String(matibayako!nameofvendor))
            End If
            If Null2String(matibayako!nameofvendor) <> "999999" Then
                gconDMIS.Execute ("Update all_Vendor set Code = '" & new_customer_code & "' where ID ='" & matibayako!ID & "'")
                Screen.MousePointer = 11
                gconDMIS.Execute "delete from ALL_VenCtl"
                For k = 65 To 90
                    Set rsCustomer = New ADODB.Recordset
                    rsCustomer.Open "select Code from ALL_Vendor where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
                    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                        NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!Code, 2, 5)) + 1, "00000")
                        gconDMIS.Execute "insert into ALL_VenCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Vendor control character for " & Chr(k) & " -')"
                    Else
                        gconDMIS.Execute "insert into ALL_VenCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Vendor control character for " & Chr(k) & " -')"
                    End If
                Next
                Screen.MousePointer = 0
            End If
            matibayako.MoveNext
            DoEvents
        Loop
    End If
    MsgBox "Ok"
End Sub

Private Sub cmdReUpdateTranType_Click()
    Screen.MousePointer = 11
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Set RSTDAYTRAN = New ADODB.Recordset
    Set RSTDAYTRAN = gconDMIS.Execute("Select ID,STOCK_ORD,TYPE from PMIS_Daytran order by id asc")
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            gconDMIS.Execute ("Update PMIS_Daytran SET TYPE = '" & SetTYPE(RSTDAYTRAN!STOCK_ORD) & "' where ID = " & RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    Set RSTDAYTRAN = New ADODB.Recordset
    Set RSTDAYTRAN = gconDMIS.Execute("Select ID,STOCK_ORD,TYPE from PMIS_TDaytran order by id asc")
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            gconDMIS.Execute ("Update PMIS_TDaytran SET TYPE = '" & SetTYPE(RSTDAYTRAN!STOCK_ORD) & "' where ID = " & RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "TAPOS"
End Sub

Private Sub cmdSetCustCodeInCusVeh_Click()
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer_Table Order by CUSCDE asc")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        rsCustomer.MoveFirst
        Do While Not rsCustomer.EOF
            gconDMIS.Execute ("Update CSMS_CUSVEH SET CUSCDE = " & N2Str2Null(rsCustomer!CUSCDE) & " where PRODNO = " & N2Str2Null(rsCustomer!ACCOUNTNO))
            rsCustomer.MoveNext
        Loop
    End If
End Sub

Private Sub cmdSetGenuineNonGenuine_Click()
    Dim rsDNPP                                         As ADODB.Recordset
    gconDMIS.Execute ("UPDATE PMIS_TDAYTRAN SET NON_HARI = 'Y' WHERE NON_HARI <> 'O'")
    gconDMIS.Execute ("UPDATE PMIS_DAYTRAN SET NON_HARI = 'Y' WHERE NON_HARI <> 'O'")
    gconDMIS.Execute ("UPDATE PMIS_STOCKMAS SET NON_HARI = 'Y',GENUINE = 'N' WHERE NON_HARI <> 'O'")

    Set rsDNPP = New ADODB.Recordset
    Set rsDNPP = gconDMIS.Execute("Select * from PMIS_DNPP Order by PARTNUMBER ASC")

    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        rsDNPP.MoveFirst
        Do While Not rsDNPP.EOF
            gconDMIS.Execute ("UPDATE PMIS_TDAYTRAN SET NON_HARI = 'N' WHERE STOCK_ORD = " & N2Str2Null(rsDNPP!PARTNUMBER))
            gconDMIS.Execute ("UPDATE PMIS_DAYTRAN SET NON_HARI = 'N' WHERE STOCK_ORD = " & N2Str2Null(rsDNPP!PARTNUMBER))
            gconDMIS.Execute ("UPDATE PMIS_STOCKMAS SET NON_HARI = 'N',GENUINE = 'Y' WHERE STOCKNO = " & N2Str2Null(rsDNPP!PARTNUMBER))
            rsDNPP.MoveNext
        Loop
    End If
    Set rsDNPP = Nothing
End Sub

Private Sub cmdSetStockMasHARINONHARI_Click()
    Dim rsSTOCKMAS                                     As ADODB.Recordset
    Dim rsSTKSTAT                                      As ADODB.Recordset
    Screen.MousePointer = 11

    Set rsSTKSTAT = New ADODB.Recordset
    Set rsSTKSTAT = gconDMIS.Execute("SELECT * FROM PMIS_STKSTAT order by STOCKNO ASC")
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        rsSTKSTAT.MoveFirst
        Do While Not rsSTKSTAT.EOF
            Me.Caption = rsSTKSTAT!STOCKNO: DoEvents
            Set rsSTOCKMAS = New ADODB.Recordset
            Set rsSTOCKMAS = gconDMIS.Execute("Select STOCKTYPE, NON_HARI from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsSTKSTAT!STOCKNO))
            If Not rsSTOCKMAS.EOF And Not rsSTOCKMAS.BOF Then
                gconDMIS.Execute ("UPDATE PMIS_STKSTAT SET STOCKTYPE = " & N2Str2Null(rsSTOCKMAS!StockType) & ", NON_HARI = " & N2Str2Null(rsSTOCKMAS!NON_HARI) & " WHERE STOCKNO = " & N2Str2Null(rsSTKSTAT!STOCKNO))
            End If
            rsSTKSTAT.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Set RSTDAYTRAN = New ADODB.Recordset
    Set RSTDAYTRAN = gconDMIS.Execute("Select ID,STOCK_ORD from PMIS_Daytran order by id asc")
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            gconDMIS.Execute ("Update PMIS_Daytran SET NON_HARI = '" & SetHariOrNonHari(RSTDAYTRAN!STOCK_ORD) & "' where ID = " & RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    Set RSTDAYTRAN = New ADODB.Recordset
    Set RSTDAYTRAN = gconDMIS.Execute("Select ID,STOCK_ORD from PMIS_TDaytran order by id asc")
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            gconDMIS.Execute ("Update PMIS_TDaytran SET NON_HARI = '" & SetHariOrNonHari(RSTDAYTRAN!STOCK_ORD) & "' where ID = " & RSTDAYTRAN!ID)
            RSTDAYTRAN.MoveNext
        Loop
    End If
    MsgBox "TAPOS NA"
End Sub

Private Sub cmdStorePartsBegBal_Click()
    Dim rsSTOCKMAS                                     As ADODB.Recordset
    Set rsSTOCKMAS = New ADODB.Recordset
    Dim CNT                                            As Integer
    Set rsSTOCKMAS = gconDMIS.Execute("Select * from PMIS_StockMas where lastm_oh > 0 Order by Id Asc")
    If Not rsSTOCKMAS.EOF And Not rsSTOCKMAS.BOF Then
        gconDMIS.Execute ("Delete from PMIS_daytran where trantype = 'BEG'")
        rsSTOCKMAS.MoveFirst: CNT = 0
        Do While Not rsSTOCKMAS.EOF
            CNT = CNT + 1
            gconDMIS.Execute ("Insert into PMIS_dayTran (ID,TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,STATUS,IN_OUT,MAC,TRANINVAMT,NON_HARI)" & _
                            " values (" & CNT & "," & N2Str2Null(rsSTOCKMAS!Type) & "," & N2Str2Null(DateSerial(Year(firstDay(LOGDATE)), Month(firstDay(LOGDATE)), Day(firstDay(LOGDATE))) - 1) & ",'BEG','111111','1111'," & N2Str2Null(rsSTOCKMAS!STOCKNO) & "," & N2Str2Null(rsSTOCKMAS!STOCKNO) & "," & N2Str2Zero(rsSTOCKMAS!LASTM_OH) & "," & N2Str2Zero(rsSTOCKMAS!MAC) & ",'P','I'," & N2Str2Zero(rsSTOCKMAS!MAC) & "," & N2Str2Zero(rsSTOCKMAS!MAC) * N2Str2Zero(rsSTOCKMAS!ONHAND) & "," & N2Str2Zero(rsSTOCKMAS!NON_HARI) & ")")
            gconDMIS.Execute ("Update PMIS_StockMas set DATE_ENTERED = " & N2Str2Null(DateSerial(Year(firstDay(LOGDATE)), Month(firstDay(LOGDATE)), Day(firstDay(LOGDATE))) - 1) & " WHERE DATE_ENTERED IS NULL AND ID = " & rsSTOCKMAS!ID)
            gconDMIS.Execute "insert into PMIS_StkStat " & _
                             "(TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,onhand,mac)" & _
                           " select TYPE, STOCKTYPE, NON_HARI, STOCKNO,STOCKDESC,LASTM_OH,LASTM_Mac from PMIS_STOCKMAS WHERE ID = " & rsSTOCKMAS!ID
            frmMain.Caption = Null2String(rsSTOCKMAS!STOCKNO) & " -> TYPE = " & Null2String(rsSTOCKMAS!Type): DoEvents
            rsSTOCKMAS.MoveNext
        Loop
    End If
    gconDMIS.Execute "update PMIS_StkStat set date_gen = " & N2Str2Null(DateSerial(Year(firstDay(LOGDATE)), Month(firstDay(LOGDATE)), Day(firstDay(LOGDATE))) - 1) & " where date_gen IS NULL"
    MsgBox "Completed"
End Sub

Private Sub cmdUploadMaterials_Click()
    Dim RSUPLOAD                                       As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset

    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select * from Z_UPLOAD_MATERIALS order by stockno asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        RSUPLOAD.MoveFirst
        Do While Not RSUPLOAD.EOF
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_STOCKMAS where stockno = " & N2Str2Null(RSUPLOAD!STOCKNO))
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                gconDMIS.Execute ("UPDATE pmis_Stockmas set " & _
                                " TYPE='M', onhand = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                " MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                " DNP = " & N2Str2Zero(RSUPLOAD!dnp) & "," & _
                                " LASTM_MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                " LASTM_OH = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                " stockdesc = " & N2Str2Null(RSUPLOAD!Description) & "," & _
                                " STOCKTYPE = " & N2Str2Null(RSUPLOAD!UOM) & "," & _
                                " SRP = " & N2Str2Null(RSUPLOAD!SRP) & "," & _
                                " active = 'Y'" & _
                                " where stockno = " & N2Str2Null(RSUPLOAD!STOCKNO))
            Else
                gconDMIS.Execute ("Insert into PMIS_Stockmas " & _
                                  "(TYPE,onhand,stockno,stockdesc,srp,mac,dnp,lastm_mac,lastm_oh,stocktype,active)" & _
                                " values ('M'," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(RSUPLOAD!STOCKNO) & "," & N2Str2Null(RSUPLOAD!Description) & "," & N2Str2Zero(RSUPLOAD!SRP) & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!MAC) * 1.12 & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(Trim(RSUPLOAD!UOM)) & ",'Y')")
            End If
            frmMain.Caption = Null2String(RSUPLOAD!STOCKNO): DoEvents
            RSUPLOAD.MoveNext
        Loop
    End If
End Sub

Private Sub cmdUploadAccessories_Click()
    Dim RSUPLOAD                                       As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset

    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select * from Z_UPLOAD_ACCESSORIES order by stockno asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        RSUPLOAD.MoveFirst
        Do While Not RSUPLOAD.EOF
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_STOCKMAS where stockno = " & N2Str2Null(RSUPLOAD!STOCKNO))
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                gconDMIS.Execute ("UPDATE pmis_Stockmas set " & _
                                " TYPE='A', onhand = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                " MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                " DNP = " & N2Str2Zero(RSUPLOAD!dnp) & "," & _
                                " LASTM_MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                " LASTM_OH = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                " stockdesc = " & N2Str2Null(RSUPLOAD!Description) & "," & _
                                " STOCKTYPE = " & N2Str2Null(RSUPLOAD!UOM) & "," & _
                                " SRP = " & N2Str2Null(RSUPLOAD!SRP) & "," & _
                                " active = 'Y'" & _
                                " where stockno = " & N2Str2Null(RSUPLOAD!STOCKNO))
            Else
                gconDMIS.Execute ("Insert into PMIS_Stockmas " & _
                                  "(TYPE,onhand,stockno,stockdesc,srp,mac,dnp,lastm_mac,lastm_oh,stocktype,active)" & _
                                " values ('A'," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(RSUPLOAD!STOCKNO) & "," & N2Str2Null(RSUPLOAD!Description) & "," & N2Str2Zero(RSUPLOAD!SRP) & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!MAC) * 1.12 & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(Trim(RSUPLOAD!UOM)) & ",'Y')")
            End If
            frmMain.Caption = Null2String(RSUPLOAD!STOCKNO): DoEvents
            RSUPLOAD.MoveNext
        Loop
    End If
End Sub

Private Sub cmdUploadParts_Click()
    Dim RSUPLOAD                                       As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset

    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select * from wizweirdo.Z_UPLOAD_PARTS order by partno asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        RSUPLOAD.MoveFirst
        Do While Not RSUPLOAD.EOF
            Set RSPARTMAS = New ADODB.Recordset
            Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_STOCKMAS where stockno = " & N2Str2Null(RSUPLOAD!PARTNO))
            If Not (RSPARTMAS.EOF Or RSPARTMAS.BOF) Then
                Call gconDMIS.Execute("UPDATE pmis_Stockmas set " & _
                                    " TYPE='P', onhand = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                    " MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                    " DNP = " & N2Str2Zero(RSUPLOAD!dnp) & "," & _
                                    " LASTM_MAC = " & N2Str2Zero(RSUPLOAD!MAC) & "," & _
                                    " LASTM_OH = " & N2Str2Zero(RSUPLOAD!oh) & "," & _
                                    " LOCATION = " & N2Str2Zero(RSUPLOAD!Location) & "," & _
                                    " SRP = " & N2Str2Zero(RSUPLOAD!SRP) & "," & _
                                    " stockdesc = " & N2Str2Null(RSUPLOAD!PARTDESC) & "," & _
                                    " STOCKTYPE = " & N2Str2Null(RSUPLOAD!Group) & "," & _
                                    " active = 'Y'" & _
                                    " where UPPER(stockno) = " & UCase(N2Str2Null(RSUPLOAD!PARTNO)))


            Else

                gconDMIS.Execute ("Insert into PMIS_Stockmas " & _
                                  "(TYPE, onhand,stockno,stockdesc,SRP,mac,dnp,lastm_mac,lastm_oh,location,stocktype,active)" & _
                                " values ('P'," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(RSUPLOAD!PARTNO) & "," & N2Str2Null(RSUPLOAD!PARTDESC) & "," & N2Str2Zero(RSUPLOAD!SRP) & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!dnp) & "," & N2Str2Zero(RSUPLOAD!MAC) & "," & N2Str2Zero(RSUPLOAD!oh) & "," & N2Str2Null(RSUPLOAD!Location) & "," & N2Str2Null(RSUPLOAD!Group) & ",'Y')")
            End If

            frmMain.Caption = Null2String(RSUPLOAD!PARTNO): DoEvents
            RSUPLOAD.MoveNext
        Loop
    End If


End Sub

