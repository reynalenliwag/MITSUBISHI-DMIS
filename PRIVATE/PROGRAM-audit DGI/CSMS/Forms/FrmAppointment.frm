VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14130
   StartUpPosition =   1  'CenterOwner
   Begin FlexCell.Grid grdApointment 
      Height          =   6915
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   12197
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   30
   End
   Begin Crystal.CrystalReport rptNARD 
      Left            =   930
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
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
         StartOfWeek     =   48037889
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
            Size            =   9.01
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
      Begin VB.Menu mnuPrintAppointment 
         Caption         =   "&Print Appointment"
      End
   End
End
Attribute VB_Name = "frmCSMSAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveorEdit                                         As String
Dim CTL                                                As Control
Dim x                                                  As Long
Dim showchange                                         As Boolean
Dim theTranno As String

Function GetNewROno(XXX As Variant)
    Dim rsNewRO                                        As New ADODB.Recordset
    Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc")
    If Not rsNewRO.EOF And Not rsNewRO.BOF Then
        GetNewROno = Format(NumericVal(Mid$(rsNewRO!rep_OR, 3, 8)) + 1, "00000000")
    Else
        GetNewROno = "00000001"
    End If
    Set rsNewRO = Nothing
End Function

Function getMake(XXX As String)
    Dim rsGetMake                                      As New ADODB.Recordset
    rsGetMake.Open "select [make] from [s_model] where [model] = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsGetMake.EOF And Not rsGetMake.BOF Then
        getMake = rsGetMake![Make]
    End If
End Function

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
        .Column(3).Width = 310: .Column(3).Locked = True
        .Column(4).Width = 0: .Column(4).Locked = True
        .Column(5).Width = 125: .Column(5).Locked = True
        .Column(6).Width = 70: .Column(6).Locked = True
        .Column(7).Width = 65: .Column(7).Locked = True
        .Column(8).Width = 0: .Column(8).Locked = True
        .Column(9).Width = 50: .Column(9).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 9, .Rows - 1, 9).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Sub ViewAppointmentGrid()
    showchange = False: InitViewInfo

    Dim rsViewGrid                                     As New ADODB.Recordset
    rsViewGrid.Open "select [status], ApptTime, CUSCDE, CUSNAM, MAKE, MODEL, PLATE_NO, KM_RDG, ApptNo from CSMS_Appointment where trandate = '" & Format(monAppointment, "MM/dd/yyyy") & "' order by right(ApptTime,2) asc,left(ApptTime,5) asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsViewGrid.EOF And Not rsViewGrid.BOF Then
        grdApointment.Rows = 1
        Do While Not rsViewGrid.EOF
            grdApointment.AddItem rsViewGrid![APPTTIME] & vbTab & _
                rsViewGrid![CUSCDE] & vbTab & _
                rsViewGrid![CUSNAM] & vbTab & _
                rsViewGrid![Make] & vbTab & _
                rsViewGrid![MODEL] & vbTab & _
                rsViewGrid![PLATE_NO] & vbTab & _
                rsViewGrid![km_rdg] & vbTab & _
                rsViewGrid![APPTNO] & vbTab & _
                rsViewGrid![Status], False
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

Sub FillGridwithTime()
    Screen.MousePointer = 11
    Call MakeApptNo
    Dim xTranDate                                       As String
    Dim xApptTime                                       As String
    Dim xApptNo                                         As Long
    Dim rsViewGrid                                      As New ADODB.Recordset
    Dim rsappt                                          As New ADODB.Recordset
    
    rsViewGrid.Open "select  TimeInterval from CSMS_ApptSchedule order by ID asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsViewGrid.EOF And Not rsViewGrid.BOF Then
        xApptNo = NumericVal(txtApptno)
        xTranDate = N2Str2Null(Format(monAppointment, "MM/dd/yyyy"))
        Do While Not rsViewGrid.EOF
            xApptTime = N2Str2Null(rsViewGrid![timeInterval])
            Set rsappt = gconDMIS.Execute("select * from CSMS_APPOINTMENT where ApptTime='" & rsViewGrid!timeInterval & "' and trandate = " & xTranDate)
            If rsappt.EOF Or rsappt.BOF Then
                gconDMIS.Execute "Insert into CSMS_Appointment " & _
                    "(ApptNo,TranDate,ApptTime)" & _
                    " values ('" & Format(xApptNo, "000000000") & _
                    "', " & xTranDate & _
                    ", " & xApptTime & ") "
            End If
            grdApointment.AddItem rsViewGrid![timeInterval], False
            xApptNo = xApptNo + 1
            rsViewGrid.MoveNext
        Loop
    End If

    grdApointment.AutoRedraw = True
    grdApointment.Refresh

    txtApptno = ""
    Screen.MousePointer = 0
End Sub

Sub MakeApptNo()
    Dim rsMakeAptNo                                    As New ADODB.Recordset
    rsMakeAptNo.Open "select [ApptNo] from CSMS_Appointment order by ApptNo desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMakeAptNo.EOF And Not rsMakeAptNo.BOF Then
        txtApptno = Format(Val(rsMakeAptNo![APPTNO]) + 1, "000000000")
    Else
        txtApptno = Format(1, "000000000")
    End If
End Sub

Sub InitViewInfo()
    txtnote = ""
    txtApptno = ""
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
    Dim rsCusVeh                                       As New ADODB.Recordset
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim rsAppointment                                  As New ADODB.Recordset
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSUPLOAD = gconDMIS.Execute("select * from csms_repairorder where apptno = '" & XXX & "'")
    If Not RSUPLOAD.EOF Or Not RSUPLOAD.BOF Then
        txtPlateNo = Null2String(RSUPLOAD!PLATE_NO)
        cboRecd_by = Null2String(RSUPLOAD!writer)
        txtnote = Null2String(RSUPLOAD!RECOMMENDATION)
        If IsDate(RSUPLOAD!PromiseDate) = True Then
            txtDte_recd = DateValue(RSUPLOAD!PromiseDate)
        End If
        ViewJobs Null2String(RSUPLOAD!RO_NO)
    End If

    Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptNo = '" & XXX & "'")
    If Not rsAppointment.EOF And Not rsAppointment.BOF Then
        txtnote = txtnote & " " & Null2String(rsAppointment!NOTE)
        txtKm_rdg = Null2String(rsAppointment!km_rdg)
        txtCustCode = Null2String(rsAppointment!CUSCDE)
        txtCustName = Null2String(rsAppointment!CUSNAM)
    End If

    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CUSVEH where PLATE_NO = '" & txtPlateNo & "' AND CUSCDE=" & N2Str2Null(txtCustCode))
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        txtPlateNo = Null2String(rsCusVeh!PLATE_NO)
        txtModel = Null2String(rsCusVeh!MODEL)
        txtMake = Null2String(rsCusVeh!Make)
        txtDescription = Null2String(rsCusVeh!Description)
        txtVIN = UCase(Null2String(rsCusVeh!VIN))
    End If

    Set RSTMP = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & txtCustCode & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblCN2 = Null2String(RSTMP!HomePhone)
        lblCN1 = Null2String(RSTMP!TelephoneNo) & " " & Null2String(RSTMP!Mobile)
    End If

    Set RSTMP = Nothing
End Sub

Sub ViewJobs(zRONO As String)
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    'JOBS
    lstJob4Service.Sorted = False: lstJob4Service.ListItems.Clear
    Set RSUPLOAD = gconDMIS.Execute("Select JOBTYPE, upper(DETCDE),DETAIL ,HRSWRK  from CSMS_Ro_Det where LIVIL='1' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Call Listview_Loadval(Me.lstJob4Service.ListItems, RSUPLOAD)
    End If
End Sub

Private Sub Command1_Click()
    With frmCSMSAppointment
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
    End With
    SaveorEdit = ""
    Call FillGridwithTime
    Call ViewAppointmentGrid
    txtApptno = ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call ViewAppointmentGrid
End Sub

Private Sub grdApointment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOptions
    End If
End Sub

Private Sub grdApointment_Click()
    On Error GoTo ERRORCODE:
    Call InitViewInfo
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        txtApptno = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
    Else
        txtApptno = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
        Call StoreAppInfo(txtApptno)
    End If

    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub grdApointment_RowColChange(ByVal Row As Long, ByVal Col As Long)
    If showchange = False Then Exit Sub
    InitViewInfo
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        txtApptno = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
    Else
        txtApptno = grdApointment.Cell(grdApointment.ActiveCell.Row, 8).Text
        txtApptTime = grdApointment.Cell(grdApointment.ActiveCell.Row, 1).Text
        StoreAppInfo txtApptno
    End If
End Sub

Private Sub mnuDeleteAppointment_Click()
    If Function_Access(LOGID, "acess_delete", "APPOINTMENT") = False Then Exit Sub

    If txtApptno = "" Then
        ShowNoRecord
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        MsgBox "Invalid! Appointment Already served", vbInformation, "Info."
        Exit Sub
    End If
    
    If MsgBox("delete this appointment, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    'VERY WRONG LOGIC MADE! - FML 03032008
    'gconDMIS.Execute ("DELETE FROM CSMS_APPOINTMENT WHERE APPTNO='" & txtApptNo & "'")

    SQL_STATEMENT = "Update CSMS_Appointment Set CUSCDE = NULL, CUSNAM = NULL, PLATE_NO = NULL, MODEL = NULL, MAKE = NULL, KM_RDG = NULL, NOTE = NULL, STATUS = NULL Where ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_Appointment"), "", "APPT NO: " & txtApptno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Delete from CSMS_RepairOrder Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno & " - SERVICE COUNTER", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Delete from CSMS_Repor Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Delete from CSMS_RO_Det Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno & " - JOBS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowDeletedMsg
    Call ViewAppointmentGrid
End Sub

Private Sub mnuOptions1_Click()
    If Function_Access(LOGID, "acess_EDIT", "APPOINTMENT") = False Then Exit Sub
    If txtApptno = "" Then
        MsgBox "Please select appointment schedule time...", vbInformation, "CSMS"
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        MessagePop InfoFriend, "Appointment Information", "Appointment has been served!", 1000
        Exit Sub
    End If

    frmCSMSEditAppointment.StoreApptTme
    frmCSMSEditAppointment.lblOLDAPPTNO.Caption = txtApptno.Caption
    frmCSMSEditAppointment.StoreAppInfo txtApptno.Caption
    frmCSMSEditAppointment.Show 1
    ViewAppointmentGrid
End Sub

Private Sub mnuPrintAppointment_Click()
    Dim ans As String
    ' Update By : BTT
    If COMPANY_CODE = "HPI" Then
        mnuPrintAppointment.Visible = True
            If MsgBox("Are you sure do want print this Appointment?", vbQuestion + vbYesNo) = vbYes Then
                PrintSQLReport rptNard, CSMS_REPORT_PATH & "appointment.rpt", "{CSMS_appointment.apptno} = '" & txtApptno.Caption & "'", CSMS_REPORT_CONNECTION, 1
            End If
    End If
End Sub

Private Sub monAppointment_DateClick(ByVal DateClicked As Date)
    Call ViewAppointmentGrid
End Sub

Private Sub mnuOptions2_Click()
    'If Function_Access(LOGID, "acess_POST", "APPOINTMENT") = False Then Exit Sub
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text = "" Then
        ShowNoRecord
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 9).Text = "Served" Then
        'MsgBox "Appointment has been served", vbInformation, "CSMS"
        MessagePop InfoFriend, "Appointment Information", "Appointment has been served!", 1000
        Exit Sub
    End If
    With frmCSMSLoadApointmentToRO
        .txtAppt = txtApptno
        .txtAcct_No = grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text
        .txtCustomer = grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text
        .txtModel = grdApointment.Cell(grdApointment.ActiveCell.Row, 5).Text
        .txtPlanteNo = grdApointment.Cell(grdApointment.ActiveCell.Row, 6).Text
        .txtROno = GetNewROno(txtApptno)
    End With
    frmCSMSLoadApointmentToRO.Show 1
End Sub

Private Sub mnuOptions0_Click()
    If Function_Access(LOGID, "acess_ADD", "APPOINTMENT") = False Then Exit Sub
    If txtApptno = "" Then
        MsgBox "Please select appointment schedule time..."
        Exit Sub
    End If
    If grdApointment.Cell(grdApointment.ActiveCell.Row, 2).Text <> "" Then
        MsgBox "This time already schedule to: " & grdApointment.Cell(grdApointment.ActiveCell.Row, 3).Text & "", vbInformation
        Exit Sub
    End If
    frmCSMSNewAppointment.txtTranNo = txtApptno
    frmCSMSNewAppointment.labType(0) = "Appointment"
    frmCSMSNewAppointment.labType(1) = "Appointment"
    frmCSMSNewAppointment.GetDefaultTransactionType
    frmCSMSNewAppointment.Show 1
End Sub

Private Sub Form_Load()
    showchange = False
    monAppointment.Value = Now()
    
    Call InitGrid
    Call FillGridwithTime
    Call ViewAppointmentGrid
    showchange = True
    If COMPANY_CODE <> "HPI" Then
        mnuPrintAppointment.Visible = False
    End If
End Sub
