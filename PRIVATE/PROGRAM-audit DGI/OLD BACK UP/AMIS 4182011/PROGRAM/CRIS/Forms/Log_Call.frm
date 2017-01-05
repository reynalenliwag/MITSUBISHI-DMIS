VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Log_Call 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Call"
   ClientHeight    =   7095
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
   Icon            =   "Log_Call.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7500
      TabIndex        =   34
      Top             =   6180
      Width           =   7500
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2190
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   38
         Top             =   0
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "Log_Call.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "Log_Call.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Delete Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "Log_Call.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Edit Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "Log_Call.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Add Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "Log_Call.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Find a Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1080
            MouseIcon       =   "Log_Call.frx":1F5E
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":20B0
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Move to Next Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   390
            MouseIcon       =   "Log_Call.frx":2408
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":255A
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Move to Previous Record"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5970
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   35
         Top             =   0
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   755
            MouseIcon       =   "Log_Call.frx":28B9
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":2A0B
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Cancel "
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "Log_Call.frx":2D49
            MousePointer    =   99  'Custom
            Picture         =   "Log_Call.frx":2E9B
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Entry"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   46
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
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4395
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
         Height          =   3435
         Left            =   45
         TabIndex        =   17
         Top             =   930
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6059
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
         Text            =   "Log_Call.frx":31EB
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
      Height          =   4395
      Left            =   2835
      ScaleHeight     =   4395
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
         Height          =   1530
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
         Format          =   51642369
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
         Format          =   51642371
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
Attribute VB_Name = "frmCRIS_Log_Call"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                          As Long
Dim CustomerCode                        As String
Dim ENTRY_LOGID                         As Long
Dim rs                                  As ADODB.Recordset

Friend Sub AddCall(xProspID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    ProspectID = xProspID
    CustomerCode = xCustCode

End Sub
'Upating Code       : AXP-0713200714:47
Private Sub cmdAdd_Click()

    On Error GoTo Errorcode:
    ENTRY_LOGID = 0
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearchQuotaion.Enabled = False
    On Error Resume Next
    cboCallType.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TempRs                          As ADODB.Recordset
    txtEntityAddress = ""
    txtEntityContactperson = ""
    txtEntityEmail = ""
    txtEntityMobile = ""
    txtEntityName = ""
    txtEntityPhone = ""

    If xProspectID = 0 Then
        labEntityName = "CUSTOMER NAME"
        Set TempRs = gconDMIS.Execute("Select CUSTOMERNAME as [Name], CONTACTPERSON, PHONE, MOBILE, ADDRESS, EMAIL from CRIS_VW_ALLPROFILE WHERE CUSCDE=" & N2Str2Null(xCUSCODE))
    Else
        labEntityName = "PROSPECT NAME"
        Set TempRs = gconDMIS.Execute("Select ACCTNAME As [NAME], CONTACTPERSON, TELEPHONE as PHONE , MOBILE, ADDRESS , EMAIL  from CRIS_PROSPECTS WHERE PROSPECTID=" & N2Str2Null(xProspectID))
    End If

    If Not (TempRs.EOF Or TempRs.BOF) Then
        txtEntityAddress = Null2String(TempRs!Address)
        txtEntityContactperson = Null2String(TempRs!ContactPerson)
        txtEntityEmail = Null2String(TempRs!EMAIL)
        txtEntityMobile = Null2String(TempRs!Mobile)
        txtEntityName = Null2String(TempRs!Name)
        txtEntityPhone = Null2String(TempRs!Phone)
    End If
    Set TempRs = Nothing
End Sub


Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    picSearchQuotaion.Enabled = True
    ENTRY_LOGID = 0
    StoreMemvars
End Sub
Sub UpdateLog()
    Dim TSQL                            As String
    If ProspectID <= 0 Then Exit Sub
    TSQL = " DECLARE @DT DATETIME" & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DATETIMECALL) FROM CRIS_PROSPECT_CALLS  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGCALL=@DT, HITCOUNTER=1 WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET  LOGCALL=NULL  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub
'Upating Code       : AXP-0713200714:47
Private Sub cmdDelete_Click()


    On Error GoTo Errorcode:

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Prospect_Calls where Logid=" & ENTRY_LOGID
        UpdateLog
        FillSearchGrid txtSearch
        rsRefresh
        StoreMemvars


    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200714:47
Private Sub cmdEdit_Click()

    On Error GoTo Errorcode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearchQuotaion.Enabled = False
    On Error Resume Next
    cboCallType.SetFocus
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    picSearchQuotaion.Enabled = True
    On Error Resume Next
    txtSearch.SetFocus
End Sub

'Upating Code       : AXP-0713200714:47
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200714:47
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200714:46
Private Sub cmdSave_Click()
    Dim t1                              As String
    Dim TempRs                          As ADODB.Recordset
    Dim sql                             As String

    On Error GoTo Errorcode:

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
        sql = "INSERT INTO CRIS_Prospect_Calls "
        sql = sql & " (ProspectID,  DateTimeCall, Duration, Subject, Comments,Bound,CalledBy, CSCDE , PhoneNo) "
        sql = sql & " VALUES("
        sql = sql & ProspectID & ","
        sql = sql & t1 & ","
        sql = sql & NumericVal(txtDuration) & ","
        sql = sql & N2Str2Null(txtSubject) & ","
        sql = sql & N2Str2Null(txtComments) & ","
        sql = sql & N2Str2Null(cboCallType) & ","
        sql = sql & N2Str2Null(txtCalledBy) & ","
        sql = sql & N2Str2Null(CustomerCode) & ","
        sql = sql & N2Str2Null(txtPhoneNo) & ")"

    Else


        sql = "Update CRIS_Prospect_Calls SET "
        sql = sql & " ProspectID=" & ProspectID & ", "
        sql = sql & " DateTimeCall=" & t1 & ", "
        sql = sql & " Duration=" & NumericVal(txtDuration) & ", "
        sql = sql & " Subject=" & N2Str2Null(txtSubject) & ", "
        sql = sql & " Comments=" & N2Str2Null(txtComments) & ", "
        sql = sql & " CalledBy =" & N2Str2Null(txtCalledBy) & ", "
        sql = sql & " PhoneNo =" & N2Str2Null(txtPhoneNo) & ", "
        sql = sql & " CSCDE =" & N2Str2Null(CustomerCode) & ", "
        sql = sql & " Bound=" & N2Str2Null(cboCallType)
        sql = sql & " WHERE LogID=" & ENTRY_LOGID
    End If

    gconDMIS.Execute (sql)
    If ENTRY_LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Phone Log Sucessfully Added", 1000
    Else
        MessagePop RecSaveOk, "RecordSaved", "Phone Log Sucessfully Updated", 1000
    End If

    UpdateLog
    rs.Requery
    If ENTRY_LOGID > 0 Then
        rs.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSearch
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Sub FillSearchGrid(XXX As String)
    Dim TempRs                          As ADODB.Recordset
    ListView1.Enabled = False
    If optAcctName.Value = True Then
        If CustomerCode <> vbNullString Then
            Set TempRs = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where  CSCDE=" & N2Str2Null(CustomerCode) & " AND  Convert(varchar, DateTimeCall , 101)  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
        Else
            Set TempRs = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where  ProspectID=" & ProspectID & " AND  Convert(varchar, DateTimeCall , 101)  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
        End If
    Else
        If CustomerCode <> vbNullString Then
            Set TempRs = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where CSCDE=" & N2Str2Null(CustomerCode) & " AND  CalledBy like '" & ReplaceQuote(XXX) & "%' order by 1 asc")
        Else
            Set TempRs = gconDMIS.Execute("select Convert(varchar, DateTimeCall , 101) , CalledBy  , LogID from CRIS_Prospect_Calls where ProspectID=" & ProspectID & " AND  CalledBy like '" & ReplaceQuote(XXX) & "%' order by 1 asc")
        End If
    End If

    ListView1.Enabled = True
    flex_FillListView TempRs, ListView1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitData
    initMemvars
    rsRefresh
    StoreMemvars
    SetEntityDetails ProspectID, CustomerCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    ENTRY_LOGID = 0
End Sub

Sub InitData()
    With cboCallType
        .AddItem ("In Bound")
        .AddItem ("Out Bound")
        .ListIndex = 0
    End With
    txtDateCall.Value = Now
    txtTimeCall.Value = Now

    picDataEntry.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    picSearchQuotaion.Enabled = True


    AddColumnHeader "Date , Person", ListView1
    ResizeColumnHeader ListView1, "55,40"
    FillSearchGrid ""
End Sub

Sub initMemvars()
    txtCalledBy = ""
    txtComments = ""
    txtDateCall = Now
    txtTimeCall = Now
    txtDuration = 0
    txtPhoneNo = ""
    txtSubject = ""
    cboCallType.ListIndex = 0
End Sub

Sub rsRefresh()
    Set rs = New ADODB.Recordset
    If CustomerCode <> vbNullString Then
        rs.Open "SELECT * From CRIS_Prospect_Calls Where CSCDE=" & N2Str2Null(CustomerCode) & " Order BY LOGID desc", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        rs.Open "SELECT * From CRIS_Prospect_Calls Where ProspectID=" & ProspectID & " Order BY LOGID DESC", gconDMIS, adOpenKeyset, adLockReadOnly
    End If
End Sub

Sub StoreMemvars()
    If Not rs.EOF And Not rs.BOF Then
        'LogID, ProspectID, DateTimeCall, Duration, Subject, Comments, Bound, CalledBy, PhoneNo
        ENTRY_LOGID = rs!LOGID
        ProspectID = rs!ProspectID
        txtCalledBy = Null2String(rs!CalledBy)
        txtComments.Text = Null2String(rs!Comments)
        txtDateCall = FormatDateTime(rs!DateTimeCall, vbShortDate)
        txtDuration = NumericVal(rs!Duration)
        txtSubject = Null2String(rs!Subject)
        txtTimeCall = FormatDateTime(rs!DateTimeCall, vbLongTime)
        txtPhoneNo = Null2String(rs!PhoneNO)
        cboCallType = Null2String(rs!Bound)

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub


Private Sub LISTVIEW1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
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

Private Sub LISTVIEW1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rs.MoveFirst
    rs.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemvars
End Sub

Private Sub optAcctName_Click()
    FillSearchGrid txtSearch
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optDate_Click()
    FillSearchGrid txtSearch
    On Error Resume Next
    txtSearch.SetFocus
End Sub



Private Sub txtSearch_Change()
    FillSearchGrid txtSearch
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

