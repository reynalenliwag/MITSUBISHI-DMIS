VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_Letter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Letter"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogLetter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   7575
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7575
      TabIndex        =   31
      Top             =   0
      Width           =   7575
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
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   210
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
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   720
         Width           =   4935
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
         TabIndex        =   35
         Text            =   "LogLetter.frx":08CA
         Top             =   1230
         Width           =   4935
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
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   210
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
         TabIndex        =   33
         Text            =   "09175041620"
         Top             =   720
         Width           =   2370
      End
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
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1260
         Width           =   2370
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
         TabIndex        =   43
         Top             =   0
         Width           =   1410
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
         TabIndex        =   42
         Top             =   1020
         Width           =   735
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
         TabIndex        =   41
         Top             =   510
         Width           =   1470
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
         TabIndex        =   40
         Top             =   0
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
         TabIndex        =   39
         Top             =   1020
         Width           =   1230
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
         TabIndex        =   38
         Top             =   510
         Width           =   1230
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   6765
         Y1              =   1710
         Y2              =   1710
      End
   End
   Begin VB.PictureBox picDataEntry 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   2955
      ScaleHeight     =   4305
      ScaleWidth      =   4695
      TabIndex        =   5
      Top             =   1785
      Width           =   4695
      Begin VB.TextBox txtLetterSubject 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   75
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1965
         Width           =   4425
      End
      Begin VB.TextBox txtLetterFrom 
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
         Left            =   75
         TabIndex        =   11
         Top             =   885
         Width           =   4275
      End
      Begin VB.TextBox txtLetterBody 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   45
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2865
         Width           =   4425
      End
      Begin VB.TextBox txtLetterTo 
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
         Left            =   75
         TabIndex        =   13
         Top             =   1425
         Width           =   4395
      End
      Begin VB.ComboBox cboLetterBound 
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   322
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker txtdtLetter 
         Height          =   345
         Left            =   75
         TabIndex        =   8
         Top             =   315
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   54067201
         CurrentDate     =   39139
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Subject:"
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
         Left            =   75
         TabIndex        =   14
         Top             =   1725
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   75
         TabIndex        =   6
         Top             =   75
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   75
         TabIndex        =   10
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Letter Body"
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
         Left            =   75
         TabIndex        =   16
         Top             =   2595
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   75
         TabIndex        =   12
         Top             =   1185
         Width           =   210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Letter Bound"
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
         Left            =   1725
         TabIndex        =   7
         Top             =   75
         Width           =   1095
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   0
      ScaleHeight     =   4305
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1785
      Width           =   2955
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   540
         Width           =   2880
      End
      Begin VB.OptionButton optSbyDate 
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
         TabIndex        =   1
         Top             =   45
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton optSbySubject 
         Caption         =   "Letter Subject"
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
         TabIndex        =   2
         Top             =   285
         Width           =   2265
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3405
         Left            =   60
         TabIndex        =   4
         Top             =   930
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   6006
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
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   7575
      TabIndex        =   18
      Top             =   6120
      Width           =   7575
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2190
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   23
         Top             =   15
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4560
            MouseIcon       =   "LogLetter.frx":0948
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":0A9A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3870
            MouseIcon       =   "LogLetter.frx":0E00
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":0F52
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3180
            MouseIcon       =   "LogLetter.frx":127D
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":13CF
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2490
            MouseIcon       =   "LogLetter.frx":172B
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":187D
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1800
            MouseIcon       =   "LogLetter.frx":1B90
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":1CE2
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1110
            MouseIcon       =   "LogLetter.frx":1FDC
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":212E
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   420
            MouseIcon       =   "LogLetter.frx":2486
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":25D8
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5940
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   20
         Top             =   15
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "LogLetter.frx":2937
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":2A89
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "LogLetter.frx":2DC7
            MousePointer    =   99  'Custom
            Picture         =   "LogLetter.frx":2F19
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Save this Record"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim ENTRY_LOGID                                                       As Long
Dim LetterCompany                                                     As String
Dim LetterProspect                                                    As String
Dim RS                                                                As ADODB.Recordset
Dim CustomerCode                                                      As String
Dim LOCALMODULENAME                                                   As String

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset

    ListView1.Enabled = False

    If optSbyDate.Value = True Then
        Set TEMPRS = gconDMIS.Execute("SELECT DateLetter,Subject , LOGID FROM CRIS_Prospect_Letter " & _
                                    " WHERE  ProspectID=" & PROSPECTID & " AND  CONVERT(VARCHAR, DateLetter, 101)  LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 1  ASC")
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT DateLetter,Subject , LOGID FROM CRIS_Prospect_Letter " & _
                                    " WHERE ProspectID=" & PROSPECTID & " AND  Subject LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 1  ASC")
    End If

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        ListView1.Enabled = True
    End If

    flex_FillListView TEMPRS, ListView1
End Sub

Sub InitData()
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(" select ACCTNAME from CRIS_PROSPECTS where ProspectID=" & PROSPECTID)
    If Not (TEMPRS.BOF Or TEMPRS.EOF) Then
        LetterProspect = Null2String(TEMPRS.Collect(0))
    End If
    LetterCompany = COMPANY_NAME


    With cboLetterBound
        .AddItem ("IN BOUND")
        .AddItem ("OUT BOUND")
        .ListIndex = 0
    End With

    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False

    AddColumnHeader "Date , Subject", ListView1
    ResizeColumnHeader ListView1, "40,55"
    FillSearchGrid ""


End Sub

Sub InitMemVars()
    txtdtLetter = DateValue(LOGDATE)
    txtLetterBody = ""
    txtLetterFrom = ""
    txtLetterTo = ""
    txtLetterSubject = ""

End Sub

Sub rsRefresh()

    Set RS = New ADODB.Recordset
    RS.Open "SELECT * From CRIS_Prospect_Letter Where ProspectID=" & PROSPECTID & " Order BY DateLetter desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        'SELECT LogID, ProspectID, DateEmail, EmailFrom, EmailTO, Subject, Body, Bound FROM DMIS.dbo.CRIS_Prospect_Email
        ENTRY_LOGID = RS!LOGID
        PROSPECTID = RS!PROSPECTID
        txtdtLetter = DateValue(RS!DateLetter)
        txtLetterBody = Null2String(RS!Body)
        txtLetterFrom = Null2String(RS!LetterFrom)
        txtLetterTo = Null2String(RS!LetterTo)
        txtLetterSubject = Null2String(RS!Subject)

        cboLetterBound = Null2String(RS!Bound)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DateLetter) FROM CRIS_Prospect_Letter  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGLETTER=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGLETTER=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

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
    End If
    Set TEMPRS = Nothing
End Sub

Friend Sub AddLetter(xProsID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    PROSPECTID = xProsID
    CustomerCode = xCustCode
    If xProsID = 0 Then
        LOCALMODULENAME = "CUSTOMER LOG"
    Else
        LOCALMODULENAME = "PROSPECT LOG"
    End If
End Sub

Private Sub cboLetterBound_Click()
    If ENTRY_LOGID = 0 Then
        If cboLetterBound.ListIndex = 0 Then
            txtLetterFrom.Text = LetterProspect
            txtLetterTo.Text = LetterCompany
        Else
            txtLetterFrom.Text = LetterCompany
            txtLetterTo.Text = LetterProspect
        End If
    End If
End Sub

Private Sub cmdAdd_Click()

    If Function_Access(LOGID, "Acess_add", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:

    ENTRY_LOGID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    cboLetterBound.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    ENTRY_LOGID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", LOCALMODULENAME) = False Then: Exit Sub
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        'gconDMIS.Execute "delete from CRIS_Prospect_Letter  where Logid=" & ENTRY_LOGID
        SQL_STATEMENT = "delete from CRIS_Prospect_Letter  where Logid=" & ENTRY_LOGID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "LOG PROSPECT  LETTER", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        LogAudit "X", LOCALMODULENAME & " FROM  :" & txtLetterFrom & " TO " & txtLetterTo & " DATE" & txtdtLetter
        UpdateLog
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
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
    PICSEARCH.Enabled = False
    On Error Resume Next
    cboLetterBound.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
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
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim SQL                                                           As String

    On Error GoTo ErrorCode:

    If ENTRY_LOGID <= 0 Then
        SQL = " INSERT INTO CRIS_PROSPECT_LETTER ( Bound, ProspectID, DateLetter, LetterFrom, LetterTo, Subject, CSCDE ,Body) "
        SQL = SQL & " VALUES( "
        SQL = SQL & N2Str2Null(cboLetterBound) & " ,"
        SQL = SQL & PROSPECTID & " ,"
        SQL = SQL & N2Str2Null(txtdtLetter.Value) & " ,"
        SQL = SQL & N2Str2Null(txtLetterFrom) & " ,"
        SQL = SQL & N2Str2Null(txtLetterTo) & " ,"
        SQL = SQL & N2Str2Null(txtLetterSubject) & " ,"
        SQL = SQL & N2Str2Null(CustomerCode) & " ,"
        SQL = SQL & N2Str2Null(txtLetterBody) & " )"
        gconDMIS.Execute (SQL)

        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "LOG PROSPECT  LETTER", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        LogAudit "A", LOCALMODULENAME & " FROM  :" & txtLetterFrom & " TO " & txtLetterTo & " DATE" & txtdtLetter
    Else
        SQL = " Update CRIS_PROSPECT_LETTER SET "
        SQL = SQL & " Bound= " & N2Str2Null(cboLetterBound) & ","
        SQL = SQL & " DateLetter= " & N2Str2Null(txtdtLetter.Value) & ","
        SQL = SQL & " LetterFrom= " & N2Str2Null(txtdtLetter.Value) & ","
        SQL = SQL & " LetterTo= " & N2Str2Null(txtLetterTo) & ","
        SQL = SQL & " Subject= " & N2Str2Null(txtLetterSubject) & ","
        SQL = SQL & " Body= " & N2Str2Null(txtLetterBody)
        SQL = SQL & " WHERE LOGID= " & ENTRY_LOGID
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL

        NEW_LogAudit "E", "LOG PROSPECT  LETTER", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        LogAudit "E", LOCALMODULENAME & " FROM  :" & txtLetterFrom & " TO " & txtLetterTo & " DATE" & txtdtLetter

    End If

    If ENTRY_LOGID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Letter Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Record Updated", 500, 1
    End If

    UpdateLog
    RS.Requery
    If ENTRY_LOGID > 0 Then
        RS.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSEARCH
    cmdCancel.Value = True
    Set TEMPRS = Nothing

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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOG PROSPECT  LETTER)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "LOG PROSPECT  LETTER")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    InitData
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
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub optSbyDate_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optSbySubject_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH
End Sub

