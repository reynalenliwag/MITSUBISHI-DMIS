VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Log_Letter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Letter"
   ClientHeight    =   7095
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
   Icon            =   "Log_Letter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   7575
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7575
      TabIndex        =   31
      Top             =   6180
      Width           =   7575
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2190
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   35
         Top             =   0
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "Log_Letter.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "Log_Letter.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Delete Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "Log_Letter.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Edit Selected Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "Log_Letter.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Add Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "Log_Letter.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Find a Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1080
            MouseIcon       =   "Log_Letter.frx":1F5E
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":20B0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Move to Next Record"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   390
            MouseIcon       =   "Log_Letter.frx":2408
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":255A
            Style           =   1  'Graphical
            TabIndex        =   42
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
         TabIndex        =   32
         Top             =   0
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   755
            MouseIcon       =   "Log_Letter.frx":28B9
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":2A0B
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Cancel"
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "Log_Letter.frx":2D49
            MousePointer    =   99  'Custom
            Picture         =   "Log_Letter.frx":2E9B
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Save Entry"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7575
      TabIndex        =   18
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Text            =   "Log_Letter.frx":31EB
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   2955
      ScaleHeight     =   4395
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
         Height          =   1515
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
         Format          =   51707905
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
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   0
      ScaleHeight     =   4395
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
         Height          =   3495
         Left            =   60
         TabIndex        =   4
         Top             =   930
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   6165
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
End
Attribute VB_Name = "frmCRIS_Log_Letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                          As Long
Dim ENTRY_LOGID                         As Long
Dim LetterCompany                       As String
Dim LetterProspect                      As String
Dim rs                                  As ADODB.Recordset
Dim CustomerCode                        As String

Friend Sub AddLetter(xProsID As Long, xCustCode As String)
    ENTRY_LOGID = 0
    ProspectID = xProsID
    CustomerCode = xCustCode
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

'Upating Code       : AXP-0713200714:48
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:

    ENTRY_LOGID = 0
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    On Error Resume Next
    cboLetterBound.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200714:48
Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:

    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    picSearch.Enabled = True
    ENTRY_LOGID = 0
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from CRIS_Prospect_Letter  where Logid=" & ENTRY_LOGID
        UpdateLog
        FillSearchGrid txtSearch
        rsRefresh
        StoreMemvars


    End If
End Sub

'Upating Code       : AXP-0713200714:48
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    picSearch.Enabled = False
    On Error Resume Next
    cboLetterBound.SetFocus
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:48
Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus

End Sub

'Upating Code       : AXP-0713200714:48
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

'Upating Code       : AXP-0713200714:48
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

'Upating Code       : AXP-0713200714:48
Private Sub cmdSave_Click()
    Dim TempRs                          As ADODB.Recordset
    Dim sql                             As String

    On Error GoTo Errorcode:

    If ENTRY_LOGID <= 0 Then
        sql = " INSERT INTO CRIS_PROSPECT_LETTER ( Bound, ProspectID, DateLetter, LetterFrom, LetterTo, Subject, CSCDE ,Body) "
        sql = sql & " VALUES( "
        sql = sql & N2Str2Null(cboLetterBound) & " ,"
        sql = sql & ProspectID & " ,"
        sql = sql & N2Str2Null(txtdtLetter.Value) & " ,"
        sql = sql & N2Str2Null(txtLetterFrom) & " ,"
        sql = sql & N2Str2Null(txtLetterTo) & " ,"
        sql = sql & N2Str2Null(txtLetterSubject) & " ,"
        sql = sql & N2Str2Null(CustomerCode) & " ,"
        sql = sql & N2Str2Null(txtLetterBody) & " )"

    Else
        sql = " Update CRIS_PROSPECT_LETTER SET "
        sql = sql & " Bound= " & N2Str2Null(cboLetterBound) & ","
        sql = sql & " DateLetter= " & N2Str2Null(txtdtLetter.Value) & ","
        sql = sql & " LetterFrom= " & N2Str2Null(txtdtLetter.Value) & ","
        sql = sql & " LetterTo= " & N2Str2Null(txtLetterTo) & ","
        sql = sql & " Subject= " & N2Str2Null(txtLetterSubject) & ","
        sql = sql & " Body= " & N2Str2Null(txtLetterBody)
        sql = sql & " WHERE LOGID= " & ENTRY_LOGID


    End If
    gconDMIS.Execute (sql)
    If ENTRY_LOGID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Letter Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Record Updated", 500, 1
    End If

    UpdateLog
    rs.Requery
    If ENTRY_LOGID > 0 Then
        rs.Find ("LOGID=" & ENTRY_LOGID)
    End If
    FillSearchGrid txtSearch
    cmdCancel.Value = True
    Set TempRs = Nothing


    Exit Sub
Errorcode:
    ShowVBError

End Sub

Sub FillSearchGrid(XXX As String)
    Dim TempRs                          As ADODB.Recordset
    ListView1.Enabled = False
    If optSbyDate.Value = True Then
        Set TempRs = gconDMIS.Execute("SELECT DateLetter,Subject , LOGID FROM CRIS_Prospect_Letter " & _
                                    " WHERE  ProspectID=" & ProspectID & " AND  CONVERT(VARCHAR, DateLetter, 101)  LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 1  ASC")

    Else
        Set TempRs = gconDMIS.Execute("SELECT DateLetter,Subject , LOGID FROM CRIS_Prospect_Letter " & _
                                    " WHERE ProspectID=" & ProspectID & " AND  Subject LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 1  ASC")
    End If
    flex_FillListView TempRs, ListView1
    ListView1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initMemvars
    InitData
    rsRefresh
    StoreMemvars
    SetEntityDetails ProspectID, CustomerCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    ENTRY_LOGID = 0
End Sub

Sub InitData()
    Dim TempRs                          As ADODB.Recordset
    Set TempRs = gconDMIS.Execute(" select ACCTNAME from CRIS_PROSPECTS where ProspectID=" & ProspectID)
    If Not (TempRs.BOF Or TempRs.EOF) Then
        LetterProspect = Null2String(TempRs.Collect(0))
    End If
    LetterCompany = COMPANY_NAME


    With cboLetterBound
        .AddItem ("IN BOUND")
        .AddItem ("OUT BOUND")
        .ListIndex = 0
    End With

    picDataEntry.Enabled = False
    picSearch.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False

    AddColumnHeader "Date , Subject", ListView1
    ResizeColumnHeader ListView1, "40,55"
    FillSearchGrid ""


End Sub

Sub initMemvars()
    txtdtLetter = DateValue(Now)
    txtLetterBody = ""
    txtLetterFrom = ""
    txtLetterTo = ""
    txtLetterSubject = ""

End Sub

Sub rsRefresh()

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * From CRIS_Prospect_Letter Where ProspectID=" & ProspectID & " Order BY DateLetter desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemvars()
    If Not rs.EOF And Not rs.BOF Then
        'SELECT LogID, ProspectID, DateEmail, EmailFrom, EmailTO, Subject, Body, Bound FROM DMIS.dbo.CRIS_Prospect_Email
        ENTRY_LOGID = rs!LOGID
        ProspectID = rs!ProspectID
        txtdtLetter = DateValue(rs!DateLetter)
        txtLetterBody = Null2String(rs!Body)
        txtLetterFrom = Null2String(rs!LetterFrom)
        txtLetterTo = Null2String(rs!LetterTo)
        txtLetterSubject = Null2String(rs!Subject)

        cboLetterBound = Null2String(rs!Bound)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateLog()
    Dim TSQL                            As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DateLetter) FROM CRIS_Prospect_Letter  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGLETTER=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGLETTER=NULL  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
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
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub
Private Sub LISTVIEW1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rs.MoveFirst
    rs.Find ("LOGID=" & Item.ListSubItems(2).Text)
    StoreMemvars
End Sub

Private Sub optSbyDate_Click()
    FillSearchGrid txtSearch
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optSbySubject_Click()
    FillSearchGrid txtSearch
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid txtSearch
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

