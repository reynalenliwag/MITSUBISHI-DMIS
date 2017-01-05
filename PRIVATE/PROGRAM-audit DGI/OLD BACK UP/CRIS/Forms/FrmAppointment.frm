VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmCSMSAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Appointment"
   ClientHeight    =   9675
   ClientLeft      =   870
   ClientTop       =   1380
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "FrmAppointment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14130
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   9675
      Left            =   10785
      ScaleHeight     =   9645
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   0
      Width           =   3350
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1740
         Picture         =   "FrmAppointment.frx":08CA
         TabIndex        =   35
         Top             =   2940
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   15
         Picture         =   "FrmAppointment.frx":0D5E
         TabIndex        =   34
         Top             =   2940
         Width           =   1725
      End
      Begin VB.TextBox txtnote 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   15
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   7080
         Width           =   3285
      End
      Begin MSComCtl2.MonthView monAppointment 
         Height          =   2610
         Left            =   15
         TabIndex        =   33
         Top             =   330
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   4604
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmAppointment.frx":1153
         MonthBackColor  =   16777215
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   51970049
         TitleBackColor  =   8388608
         TitleForeColor  =   16777215
         TrailingForeColor=   13932144
         CurrentDate     =   38458
         MaxDate         =   42369
      End
      Begin VB.Label lblCN1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   15
         TabIndex        =   32
         Top             =   9330
         Width           =   3285
      End
      Begin VB.Label Label11 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Contacts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   31
         Top             =   9030
         Width           =   1275
      End
      Begin VB.Label lblCN2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   30
         Top             =   9030
         Width           =   1995
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D2BDB6&
         Caption         =   " VIN No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   29
         Top             =   8430
         Width           =   1275
      End
      Begin VB.Label txtVIN 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   28
         Top             =   8430
         Width           =   1995
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D2BDB6&
         Caption         =   " KM Rdg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   27
         Top             =   8730
         Width           =   1275
      End
      Begin VB.Label txtKm_rdg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   26
         Top             =   8730
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Promise Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   25
         Top             =   8130
         Width           =   1275
      End
      Begin VB.Label txtDte_recd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   24
         Top             =   8130
         Width           =   1995
      End
      Begin VB.Label lblAppt 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Customer Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   23
         Top             =   3990
         Width           =   1275
      End
      Begin VB.Label lblCalls 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   22
         Top             =   5160
         Width           =   1275
      End
      Begin VB.Label lblLetters 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Make"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   21
         Top             =   5460
         Width           =   1275
      End
      Begin VB.Label lblTest 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Schedule Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   20
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label lblVisits 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Plate Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   19
         Top             =   4860
         Width           =   1275
      End
      Begin VB.Label lblLoan 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   18
         Top             =   6780
         Width           =   1275
      End
      Begin VB.Label txtCustName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   555
         Left            =   15
         TabIndex        =   17
         Top             =   4290
         Width           =   3285
      End
      Begin VB.Label lblEmails 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   16
         Top             =   5760
         Width           =   3285
      End
      Begin XtremeShortcutBar.ShortcutCaption captionInformation 
         Height          =   315
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   3345
         _Version        =   655364
         _ExtentX        =   5900
         _ExtentY        =   556
         _StockProps     =   14
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Recieved By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   14
         Top             =   7830
         Width           =   1275
      End
      Begin VB.Label txtDescription 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   705
         Left            =   15
         TabIndex        =   13
         Top             =   6060
         Width           =   3285
      End
      Begin VB.Label lblLogLoan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   12
         Top             =   6780
         Width           =   1995
      End
      Begin VB.Label txtPlateNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   11
         Top             =   4860
         Width           =   1995
      End
      Begin VB.Label txtCustCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1310
         TabIndex        =   10
         ToolTipText     =   "Last sales appointment made on and days elasped"
         Top             =   3990
         Width           =   1995
      End
      Begin VB.Label txtApptTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1310
         TabIndex        =   9
         ToolTipText     =   " Test Drive Schedules On and Day Elasped"
         Top             =   3690
         Width           =   1995
      End
      Begin VB.Label txtModel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   8
         Top             =   5160
         Width           =   1995
      End
      Begin VB.Label txtMake 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   7
         Top             =   5460
         Width           =   1995
      End
      Begin VB.Label cboRecd_by 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   6
         Top             =   7830
         Width           =   1995
      End
      Begin VB.Label lblQ 
         BackColor       =   &H00D2BDB6&
         Caption         =   " Appointment #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   15
         TabIndex        =   5
         Top             =   3390
         Width           =   1275
      End
      Begin VB.Label txtApptNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         ToolTipText     =   " Last Quotation Send "
         Top             =   3390
         Width           =   1995
      End
   End
   Begin FlexCell.Grid grdApointment 
      Height          =   7155
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   12621
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   30
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   390
      TabIndex        =   1
      Top             =   1290
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin MSComctlLib.ListView lstJob4Service 
      Height          =   2415
      Left            =   30
      TabIndex        =   36
      Top             =   7230
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmAppointment.frx":12B5
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Job Type"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Jobs Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hrs. Work"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Menu Options"
      Visible         =   0   'False
      Begin VB.Menu mnuOptions0 
         Caption         =   "&Add Customer Appointment"
      End
      Begin VB.Menu mnuOptions1 
         Caption         =   "&Edit Selected Appointment"
      End
      Begin VB.Menu mnuOptions2 
         Caption         =   "&Upload to  Repair Order (R/O)"
      End
      Begin VB.Menu mnuDeleteAppointment 
         Caption         =   "&Delete Appointment"
      End
   End
End
Attribute VB_Name = "frmCSMSAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveorEdit                          As String
Dim ctl                                 As Control
Dim x                                   As Long
Dim showchange                          As Boolean

Private Sub Command1_Click()
    With frmCSMSAppointment
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next ctl
    End With
    SaveorEdit = ""
    FillGridwithTime
    ViewAppointmentGrid
    txtApptNo = ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub grdApointment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOptions
    End If
End Sub
Private Sub grdApointment_Click()
    On Error GoTo Errorcode:
    InitViewInfo
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        txtApptNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
    Else
        txtApptNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
        StoreAppInfo txtApptNo
        
    End If

    Exit Sub
Errorcode:
    ShowVBError
End Sub


Sub InitGrid()

    lstJob4Service.ColumnHeaders(1).Width = "1124.787"
    lstJob4Service.ColumnHeaders(2).Width = "1040.315"
    lstJob4Service.ColumnHeaders(3).Width = "7304.882"
    
 
 
 
    With grdApointment
        .Cols = 10: .Rows = 2
        .DisplayFocusRect = False: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = " TIME"
        .Cell(0, 2).Text = "Account#"
        .Cell(0, 3).Text = "Customer Name"
        .Cell(0, 4).Text = "Make"
        .Cell(0, 5).Text = "Model"
        .Cell(0, 6).Text = "Plate No."
        .Cell(0, 7).Text = "KM Rdg"
        .Cell(0, 8).Text = "Appt No"
        .Cell(0, 9).Text = "Status"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:
        .Column(3).CellType = cellTextBox:
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox

        .Column(0).Width = 18
        .Column(1).Width = 55: .Column(1).Locked = True
        .Column(2).Width = 0: .Column(2).Locked = True
        .Column(3).Width = 160: .Column(3).Locked = True
        .Column(4).Width = 0: .Column(4).Locked = True
        .Column(5).Width = 100: .Column(5).Locked = True
        .Column(6).Width = 70: .Column(6).Locked = True
        .Column(7).Width = 65: .Column(7).Locked = True
        .Column(8).Width = 0: .Column(8).Locked = True
        .Column(9).Width = 50: .Column(8).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 9, .Rows - 1, 9).ForeColor = RGB(0, 0, 128)
    End With
End Sub

 
Private Sub grdApointment_RowColChange(ByVal Row As Long, ByVal Col As Long)
    If showchange = False Then Exit Sub
    InitViewInfo
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        txtApptNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
    Else
        txtApptNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
        StoreAppInfo txtApptNo
    End If
End Sub

Private Sub mnuDeleteAppointment_Click()
    If Function_Access(LOGID, "acess_delete", "APPOINTMENT") = False Then Exit Sub

    If txtApptNo = "" Then
        MsgBox "No Appointment to be edit... ", vbInformation
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        MsgBox "Invalid! Appointment has been served...", vbInformation
        Exit Sub
    End If
    If MsgBox("Are you Sure You want to delete this appointment", vbInformation + vbYesNo) = vbNo Then Exit Sub
    'VERY WRONG LOGIC MADE! - FML 03032008
    'gconDMIS.Execute ("DELETE FROM CSMS_APPOINTMENT WHERE APPTNO='" & txtApptNo & "'")
    gconDMIS.Execute ("Update CSMS_Appointment Set CUSCDE = NULL, CUSNAM = NULL, PLATE_NO = NULL, MODEL = NULL, MAKE = NULL, KM_RDG = NULL, NOTE = NULL, STATUS = NULL Where ApptNo = '" & txtApptNo & "'")
    gconDMIS.Execute ("Delete from CSMS_RepairOrder Where TransType = 'A' AND ApptNo = '" & txtApptNo & "'")
    gconDMIS.Execute ("Delete from CSMS_Repor Where TransType = 'A' AND ApptNo = '" & txtApptNo & "'")
    gconDMIS.Execute ("Delete from CSMS_RO_Det Where TransType = 'A' AND ApptNo = '" & txtApptNo & "'")
    MsgBox "Appointment sucessfully deleted...", vbInformation
    ViewAppointmentGrid
End Sub

Private Sub mnuOptions1_Click()

    If txtApptNo = "" Then
        MsgBox "Please select appointment schedule time..."
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        MsgBox "No Appointment to be edit... ", vbInformation
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        MsgBox "Appointment Already Served!", vbInformation, "Edit Not Allowed"
        Exit Sub
    End If
    frmCSMSEditAppointment.StoreApptTme
    frmCSMSEditAppointment.StoreAppInfo txtApptNo.Caption

    frmCSMSEditAppointment.Show 1
    ViewAppointmentGrid

End Sub
Private Sub monAppointment_DateClick(ByVal DateClicked As Date)

    ViewAppointmentGrid
End Sub
Private Sub mnuOptions2_Click()


    If grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text = "" Then
        MsgBox "Invalid! no appointment has been maid..."
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        MsgBox "Invalid! Appointment has been served..."
        Exit Sub
    End If
    With frmCSMSLoadApointmentToRO
        .txtAppt = txtApptNo
        .txtAcct_No = grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text
        .txtCustomer = grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text
        .txtModel = grdApointment.Cell(grdApointment.ActiveCell.Row, 5).Text
        .txtPlanteNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 6).Text
        .txtROno = GetNewROno(txtApptNo)
    End With
    frmCSMSLoadApointmentToRO.Show 1
 
 
End Sub
Function GetNewROno(XXX As Variant)
    Dim rsNewRO                         As ADODB.Recordset
    Set rsNewRO = New ADODB.Recordset
    Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc")
    If Not rsNewRO.EOF And Not rsNewRO.BOF Then
        GetNewROno = Format(NumericVal(Mid$(rsNewRO!REP_OR, 3, 8)) + 1, "R-00000000")
    Else
        GetNewROno = "R-00000001"
    End If
    Set rsNewRO = Nothing
End Function
Private Sub mnuOptions0_Click()
    If txtApptNo = "" Then
        MsgBox "Please select appointment schedule time..."
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text <> "" Then
        MsgBox "This time already schedule to: " & grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text & "", vbInformation
        Exit Sub
    End If
    frmCSMSNewAppointment.txtTranNo = txtApptNo
    frmCSMSNewAppointment.labType(0) = "Appointment"
    frmCSMSNewAppointment.labType(1) = "Appointment"
    frmCSMSNewAppointment.GetDefaultTransactionType
    frmCSMSNewAppointment.Show 1
End Sub
Sub ViewAppointmentGrid()
    showchange = False: InitViewInfo

    Dim rsViewGrid                      As ADODB.Recordset
    Set rsViewGrid = New ADODB.Recordset
    rsViewGrid.Open "select [status],ApptTime,CUSCDE,CUSNAM,MAKE,MODEL,PLATE_NO,KM_RDG,ApptNo from CSMS_Appointment where trandate = '" & Format(monAppointment, "MM/dd/yyyy") & "' order by right(ApptTime,2) asc,left(ApptTime,5) asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsViewGrid.EOF And Not rsViewGrid.BOF Then
        grdApointment.Rows = 1
        Do While Not rsViewGrid.EOF
            grdApointment.AddItem rsViewGrid![ApptTime] & vbTab & _
                                  rsViewGrid![CUSCDE] & vbTab & _
                                  rsViewGrid![cusnam] & vbTab & _
                                  rsViewGrid![Make] & vbTab & _
                                  rsViewGrid![Model] & vbTab & _
                                  rsViewGrid![Plate_no] & vbTab & _
                                  rsViewGrid![KM_RDG] & vbTab & _
                                  rsViewGrid![ApptNo] & vbTab & _
                                  rsViewGrid![status], False
            rsViewGrid.MoveNext
        Loop
        grdApointment.AutoRedraw = True

        For x = 1 To grdApointment.Rows - 1
            If grdApointment.Cell(x, 9).Text = "Served" Then
                grdApointment.Range(x, 1, x, 9).Selected
                grdApointment.Range(x, 1, x, 9).BackColor = &HE8FEFF
                grdApointment.Range(x, 1, x, 9).FontBold = True
                grdApointment.Range(x, 1, x, 9).ForeColor = &HFF&
            End If
        Next
        If grdApointment.Rows > 0 Then
            grdApointment.Cell(1, 1).SetFocus
        End If
    Else
        FillGridwithTime
    End If
    grdApointment.AutoRedraw = True
    grdApointment.Refresh
    showchange = True
End Sub









Private Sub Form_Load()
    showchange = False
    monAppointment.Value = Now()
    InitGrid
    FillGridwithTime
    ViewAppointmentGrid
    showchange = True
End Sub



Sub FillGridwithTime()

    Screen.MousePointer = 11
    MakeApptNo
    Dim xTranDate, xApptTime            As String
    Dim xApptNo                         As Long
    Dim rsViewGrid                      As ADODB.Recordset
    Set rsViewGrid = New ADODB.Recordset
    Dim rsappt                          As ADODB.Recordset
    rsViewGrid.Open "select  TimeInterval from CSMS_ApptSchedule order by ID asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsViewGrid.EOF And Not rsViewGrid.BOF Then
        xApptNo = NumericVal(txtApptNo)
        xTranDate = N2Str2Null(Format(monAppointment, "MM/dd/yyyy"))
        Do While Not rsViewGrid.EOF
            xApptTime = N2Str2Null(rsViewGrid![timeInterval])
            Set rsappt = gconDMIS.Execute("select * from CSMS_APPOINTMENT where ApptTime='" & rsViewGrid!timeInterval & "' and trandate = " & xTranDate)
            If rsappt.EOF Or rsappt.BOF Then
                gconDMIS.Execute "Insert into CSMS_Appointment " & _
                                 "(ApptNo,TranDate,ApptTime)" & _
                               " values ('" & Format(xApptNo, "000000000") & "'," & xTranDate & "," & xApptTime & ") "
            End If
            grdApointment.AddItem rsViewGrid![timeInterval], False
            xApptNo = xApptNo + 1
            rsViewGrid.MoveNext
        Loop
    End If

    grdApointment.AutoRedraw = True
    grdApointment.Refresh

    txtApptNo = ""
    Screen.MousePointer = 0
End Sub

Function getMake(XXX As String)
    Dim rsGetMake                       As ADODB.Recordset
    Set rsGetMake = New ADODB.Recordset
    rsGetMake.Open "select [make] from [s_model] where [model] = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsGetMake.EOF And Not rsGetMake.BOF Then
        getMake = rsGetMake![Make]
    End If
End Function

Sub MakeApptNo()
    Dim rsMakeAptNo                     As ADODB.Recordset
    Set rsMakeAptNo = New ADODB.Recordset
    rsMakeAptNo.Open "select [ApptNo] from CSMS_Appointment order by ApptNo desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMakeAptNo.EOF And Not rsMakeAptNo.BOF Then
        txtApptNo = Format(Val(rsMakeAptNo![ApptNo]) + 1, "000000000")
    Else
        txtApptNo = Format(1, "000000000")
    End If
End Sub
Sub InitViewInfo()
    txtnote = ""
    txtApptNo = ""
    txtApptTime = ""
    txtnote = ""
    txtCustCode = ""
    txtCustName = ""
    txtPlateNo = ""
    cboRecd_by = ""
    txtKm_rdg = ""
    txtDte_recd = ""
    txtVIN = ""
    txtPlateNo = ""
    txtModel = ""
    txtMake = ""
    txtDescription = ""
    txtDte_recd = ""
    lblCN1 = ""
    lblCN2 = ""
     
    lstJob4Service.ListItems.Clear
     
End Sub


Sub StoreAppInfo(XXX As String)
    Dim rsCusVeh                        As ADODB.Recordset
    Dim rsUpload                        As ADODB.Recordset
    Dim rsAppointment                   As ADODB.Recordset
    Dim rsTmp                           As ADODB.Recordset
    

    Set rsUpload = gconDMIS.Execute("select * from csms_repairorder where apptno = '" & XXX & "'")
    
    If Not rsUpload.EOF Or Not rsUpload.BOF Then
        txtPlateNo = Null2String(rsUpload!Plate_no)
        cboRecd_by = Null2String(rsUpload!writer)
        txtnote = Null2String(rsUpload!recommendation)
        If IsDate(rsUpload!promisedate) = True Then
            txtDte_recd = DateValue(rsUpload!promisedate)
        End If
        ViewJobs Null2String(rsUpload!RO_NO)
    End If

    Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptNo = '" & XXX & "'")
    If Not rsAppointment.EOF And Not rsAppointment.BOF Then
        txtnote = txtnote & " " & Null2String(rsAppointment!NOTE)
        txtKm_rdg = Null2String(rsAppointment!KM_RDG)
        txtCustCode = Null2String(rsAppointment!CUSCDE)
        txtCustName = Null2String(rsAppointment!cusnam)
    End If

    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CUSVEH where PLATE_NO = '" & txtPlateNo & "' AND CUSCDE=" & N2Str2Null(txtCustCode))
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        txtPlateNo = Null2String(rsCusVeh!Plate_no)
        txtModel = Null2String(rsCusVeh!Model)
        txtMake = Null2String(rsCusVeh!Make)
        txtDescription = Null2String(rsCusVeh!Description)
        txtVIN = UCase(Null2String(rsCusVeh!Vin))
    End If







    Set rsTmp = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & txtCustCode & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        lblCN2 = Null2String(rsTmp!HomePhone)
        lblCN1 = Null2String(rsTmp!TelephoneNo) & " " & Null2String(rsTmp!Mobile)
    End If
    
    Set rsTmp = Nothing
End Sub




Sub ViewJobs(zRONO As String)
    Dim rsUpload                        As ADODB.Recordset
    Dim Item                            As ListItem

    'JOBS
    lstJob4Service.Sorted = False: lstJob4Service.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select JOBTYPE, upper(DETCDE),DETAIL ,HRSWRK  from CSMS_Ro_Det where LIVIL='1' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Call Listview_Loadval(Me.lstJob4Service.ListItems, rsUpload)
    End If
 

   
 
 
  
End Sub


Private Sub SSTab1_DblClick()

End Sub
