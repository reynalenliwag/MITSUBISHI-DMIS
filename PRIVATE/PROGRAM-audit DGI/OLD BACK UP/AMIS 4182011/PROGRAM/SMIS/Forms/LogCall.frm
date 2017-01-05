VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_Call 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Call"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogCall.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   7500
      TabIndex        =   34
      Top             =   6060
      Width           =   7500
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2100
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   39
         Top             =   15
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "LogCall.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "LogCall.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "LogCall.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "LogCall.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "LogCall.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1080
            MouseIcon       =   "LogCall.frx":1F5E
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":20B0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   390
            MouseIcon       =   "LogCall.frx":2408
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":255A
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5880
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   36
         Top             =   15
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "LogCall.frx":28B9
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":2A0B
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "LogCall.frx":2D49
            MousePointer    =   99  'Custom
            Picture         =   "LogCall.frx":2E9B
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Log Call"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox picSearchQuotaion 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   2835
      TabIndex        =   13
      Top             =   1785
      Width           =   2835
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   45
         TabIndex        =   16
         Top             =   540
         Width           =   2745
      End
      Begin VB.OptionButton optAcctName 
         Caption         =   "Search By Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   14
         Top             =   45
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Person Called"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   285
         Width           =   2265
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3315
         Left            =   45
         TabIndex        =   17
         Top             =   930
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5847
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.TextBox txtEntityEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5070
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1260
         Width           =   2370
      End
      Begin VB.TextBox txtEntityMobile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   8
         Text            =   "09175041620"
         Top             =   720
         Width           =   2370
      End
      Begin VB.TextBox txtEntityPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox txtEntityAddress 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "LogCall.frx":31EB
         Top             =   1230
         Width           =   4935
      End
      Begin VB.TextBox txtEntityContactperson 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtEntityName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   210
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   6765
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   6
         Top             =   510
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   10
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   2
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CONTACT PERSON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   5
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label labEntityAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   9
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label labEntityName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   1410
      End
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   2835
      ScaleHeight     =   4275
      ScaleWidth      =   4665
      TabIndex        =   18
      Top             =   1785
      Width           =   4665
      Begin VB.TextBox txtCalledBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   21
         Top             =   240
         Width           =   2685
      End
      Begin VB.TextBox txtPhoneNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   31
         Top             =   2130
         Width           =   4005
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2760
         Width           =   4035
      End
      Begin VB.ComboBox cboCallType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2820
         TabIndex        =   22
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   29
         Top             =   1470
         Width           =   4035
      End
      Begin VB.TextBox txtDuration 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2850
         TabIndex        =   27
         Text            =   "0"
         Top             =   870
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker txtDateCall 
         Height          =   345
         Left            =   60
         TabIndex        =   25
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   54263809
         CurrentDate     =   39139
      End
      Begin MSComCtl2.DTPicker txtTimeCall 
         Height          =   345
         Left            =   1410
         TabIndex        =   26
         Top             =   870
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   54263811
         UpDown          =   -1  'True
         CurrentDate     =   39139
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Called By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   19
         Top             =   -30
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2910
         TabIndex        =   24
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   32
         Top             =   2520
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Call Bound"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2850
         TabIndex        =   20
         Top             =   -30
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date/Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -300
         TabIndex        =   23
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   28
         Top             =   1230
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Call"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim CustomerCode                                                      As String
Dim ENTRY_LOGID                                                       As Long
Dim RS                                                                As ADODB.Recordset
Dim LOCALMODULENAME                                                   As String

Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    txtEntityAddress = ""
    txtEntityContactperson = ""
    txtEntityEmail = ""
    txtEntityMobile = ""
    txtEntityName = ""
    txtEntityPhone = ""
    If xProspectID = 0 Then
        labEntityName = "CUSTOMER NAME"
        Set TEMPRS = gconDMIS.Execute("Select CUSTOMERNAME as [Name], CONTACTPERSON, PHONE, MOBILE, ADDRESS, EMAIL from CRIS_VW_ALLPROFILE WHERE CUSCDE=" & N2Str2Null(xCUSCODE))
    Else
        labEntityName = "PROSPECT NAME"
        Set TEMPRS = gconDMIS.Execute("Select ACCTNAME As [NAME], CONTACTPERSON, TELEPHONE as PHONE , MOBILE, ADDRESS , EMAIL  from CRIS_PROSPECTS WHERE PROSPECTID=" & N2Str2Null(xProspectID))
    End If
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtEntityAddress = Null2String(TEMPRS!Address)
        txtEntityContactperson = Null2String(TEMPRS!ContactPerson)
        txtEntityEmail = Null2String(TEMPRS!EMAIL)
        txtEntityMobile = Null2String(TEMPRS!Mobile)
        txtEntityName = Null2String(TEMPRS!Name)
        txtEntityPhone = Null2String(TEMPRS!Phone)
        txtPhoneNo = Null2String(TEMPRS!Phone)
        txtCalledBy = Null2String(TEMPRS!ContactPerson)
    End If
    Set TEMPRS = Nothing
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    If PROSPECTID <= 0 Then Exit Sub
    TSQL = " DECLARE @DT DATETIME" & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DATETIMECALL) FROM CRIS_PROSPECT_CALLS  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGCALL=@DT, HITCOUNTER=1 WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET  LOGCALL=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    ListView1.Enabled = False
    If optAcctName.Value = True Then
        If CustomerCode <> vbNullString Then
            Set TEMPRS = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where  CSCDE=" & N2Str2Null(CustomerCode) & " AND  Convert(varchar, DateTimeCall , 101)  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
        Else
            Set TEMPRS = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where  ProspectID=" & PROSPECTID & " AND  Convert(varchar, DateTimeCall , 101)  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
        End If
    Else
        If CustomerCode <> vbNullString Then
            Set TEMPRS = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where CSCDE=" & N2Str2Null(CustomerCode) & " AND  CalledBy like '" & ReplaceQuote(XXX) & "%' order by 1 asc")
        Else
            Set TEMPRS = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where ProspectID=" & PROSPECTID & " AND  CalledBy like '" & ReplaceQuote(XXX) & "%' order by 1 asc")
        End If
    End If

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        ListView1.Enabled = True
    End If

    flex_FillListView TEMPRS, ListView1
End Sub

Sub InitData()
    With cboCallType
        .AddItem ("In Bound")
        .AddItem ("Out Bound")
        .ListIndex = 0
    End With
    txtDateCall.Value = LOGDATE
    txtTimeCall.Value = LOGDATE

    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    picSearchQuotaion.Enabled = True


    AddColumnHeader "Date , Person", ListView1
    ResizeColumnHeader ListView1, "55,40"
    FillSearchGrid ""
End Sub

Sub InitMemVars()
    txtCalledBy = ""
    txtComments = ""
    txtDateCall = LOGDATE
    txtTimeCall = LOGDATE
    txtDuration = 0
    txtPhoneNo = ""
    txtSubject = ""
    cboCallType.ListIndex = 0
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    If CustomerCode <> vbNullString Then
        RS.Open "SELECT * From CRIS_Prospect_Calls Where CSCDE=" & N2Str2Null(CustomerCode) & " Order BY LOGID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        RS.Open "SELECT * From CRIS_Prospect_Calls Where ProspectID=" & PROSPECTID & " Order BY LOGID DESC", gconDMIS, adOpenKeyset, adLockReadOnly
    End If
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        'LogID, ProspectID, DateTimeCall, Duration, Subject, Comments, Bound, CalledBy, PhoneNo
        ENTRY_LOGID = RS!LOGID
        PROSPECTID = RS!PROSPECTID
        txtCalledBy = Null2String(RS!CalledBy)
        txtComments.Text = Null2String(RS!Comments)
        txtDateCall = FormatDateTime(RS!DateTimeCall, vbShortDate)
        txtDuration = NumericVal(RS!Duration)
        txtSubject = Null2String(RS!Subject)
        txtTimeCall = FormatDateTime(RS!DateTimeCall, vbLongTime)
        txtPhoneNo = Null2String(RS!PhoneNO)
        cboCallType = Null2String(RS!Bound)


    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Friend Sub AddCall(xProspID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    PROSPECTID = xProspID
    CustomerCode = xCustCode
    If xProspID = 0 Then
        LOCALMODULENAME = "CUSTOMER LOG"
    Else
        LOCALMODULENAME = "PROSPECT LOG"
    End If
End Sub

Private Sub cboCallType_Click()
    If ENTRY_LOGID = 0 Then
        If cboCallType = "Out Bound" Then
            txtCalledBy = LOGSAE
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_add", LOCALMODULENAME) = False Then: picAdds.Visible = False: Exit Sub
    On Error GoTo ErrorCode:
    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearchQuotaion.Enabled = False
    On Error Resume Next
    cboCallType.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    picSearchQuotaion.Enabled = True
    ENTRY_LOGID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from CRIS_Prospect_Calls where Logid=" & ENTRY_LOGID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "LOG PROSPECT CALL", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        UpdateLog
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
        LogAudit "X", LOCALMODULENAME & " CALLED BY :" & txtCalledBy & " PHONE " & txtPhoneNo
        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearchQuotaion.Enabled = False
    On Error Resume Next
    cboCallType.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    picSearchQuotaion.Enabled = True
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Dim t1                                                            As String
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim SQL                                                           As String
    On Error GoTo ErrorCode:
    If cboCallType = "" Then
        ShowIsRequiredMsg "Call Type"
        On Error Resume Next
        cboCallType.SetFocus
        Exit Sub
    End If
    If txtSubject = "" Then
        ShowIsRequiredMsg "Subject Name "
        On Error Resume Next
        txtSubject.SetFocus
        Exit Sub
    End If
    t1 = N2Str2Null(DateValue(txtDateCall) & " " & TimeValue(txtTimeCall))
    If ENTRY_LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospect_Calls "
        SQL = SQL & " (ProspectID,  DateTimeCall, Duration, Subject, Comments,Bound,CalledBy, CSCDE , PhoneNo) "
        SQL = SQL & " VALUES("
        SQL = SQL & PROSPECTID & ","
        SQL = SQL & t1 & ","
        SQL = SQL & NumericVal(txtDuration) & ","
        SQL = SQL & N2Str2Null(txtSubject) & ","
        SQL = SQL & N2Str2Null(txtComments) & ","
        SQL = SQL & N2Str2Null(cboCallType) & ","
        SQL = SQL & N2Str2Null(txtCalledBy) & ","
        SQL = SQL & N2Str2Null(CustomerCode) & ","
        SQL = SQL & N2Str2Null(txtPhoneNo) & ")"
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "LOG PROSPECT CALL", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        LogAudit "A", LOCALMODULENAME & " CALLED BY :" & txtCalledBy & " PHONE " & txtPhoneNo
    Else
        SQL = "Update CRIS_Prospect_Calls SET "
        SQL = SQL & " ProspectID=" & PROSPECTID & ", "
        SQL = SQL & " DateTimeCall=" & t1 & ", "
        SQL = SQL & " Duration=" & NumericVal(txtDuration) & ", "
        SQL = SQL & " Subject=" & N2Str2Null(txtSubject) & ", "
        SQL = SQL & " Comments=" & N2Str2Null(txtComments) & ", "
        SQL = SQL & " CalledBy =" & N2Str2Null(txtCalledBy) & ", "
        SQL = SQL & " PhoneNo =" & N2Str2Null(txtPhoneNo) & ", "
        SQL = SQL & " CSCDE =" & N2Str2Null(CustomerCode) & ", "
        SQL = SQL & " Bound=" & N2Str2Null(cboCallType)
        SQL = SQL & " WHERE LogID=" & ENTRY_LOGID
        gconDMIS.Execute (SQL)

        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "LOG PROSPECT CALL", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""


        LogAudit "E", LOCALMODULENAME & " CALLED BY :" & txtCalledBy & " PHONE " & txtPhoneNo
    End If

    If ENTRY_LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Phone Log Sucessfully Added", 1000
    Else
        MessagePop RecSaveOk, "RecordSaved", "Phone Log Sucessfully Updated", 1000
    End If
    UpdateLog
    RS.Requery
    If ENTRY_LOGID > 0 Then
        RS.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSEARCH
    cmdCancel.Value = True
    If FormExist("MainForm") Then
        MainForm.ShowStatus PROSPECTID
    End If
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOG PROSPECT CALL)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "LOG PROSPECT CALL")
            'End If
    End Select

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
    InitMemVars
    rsRefresh
    StoreMemVars
    SetEntityDetails PROSPECTID, CustomerCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PROSPECTID = 0
    ENTRY_LOGID = 0
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub LISTVIEW1_DblClick()
    If ListView1.SelectedItem Is Nothing Then: Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub optAcctName_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optDate_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

