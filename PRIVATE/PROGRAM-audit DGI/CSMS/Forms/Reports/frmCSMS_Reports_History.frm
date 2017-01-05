VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_Reports_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer History"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   12690
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6825
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   12525
      _Version        =   655364
      _ExtentX        =   22093
      _ExtentY        =   12039
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.ComboBox cboYear 
      Height          =   330
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   210
      Width           =   1695
   End
   Begin VB.ComboBox cboMonth 
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   210
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageListReport 
      Left            =   10650
      Top             =   7170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMS_Reports_History.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMS_Reports_History.frx":0065
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMS_Reports_History.frx":00D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCSMS_Reports_History.frx":025F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   675
      Left            =   -30
      TabIndex        =   1
      Top             =   -30
      Width           =   10755
      _Version        =   655364
      _ExtentX        =   18971
      _ExtentY        =   1191
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmCSMS_Reports_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COLUMN_ICON = 0
Const COLUMN_CHECK = 1
Const COLUMN_SUBJECT = 2
Const COLUMN_FROM = 3
Const COLUMN_SENT = 4
Const COLUMN_SIZE = 5
Const COLUMN_PRICE = 6
Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call fillcbomonth(cboMonth)
    Call FillCboMoreYear(cboYear)
    Call InitGrid
End Sub

Sub InitGrid()
    Dim Column As ReportColumn
    ' Setup ReportControl columns
    Set Column = wndReportControl.Columns.Add(COLUMN_ICON, "Icon", 18, False)
    Column.Icon = 0
    Column.Sortable = False
    Set Column = wndReportControl.Columns.Add(COLUMN_CHECK, "Check", 18, False)
    Column.Icon = 1
    Column.Sortable = False
    
    Set Column = wndReportControl.Columns.Add(COLUMN_SUBJECT, "Subject", 280, True)
    Column.TreeColumn = True
    
    wndReportControl.Columns.Add COLUMN_FROM, "From", 180, True
    wndReportControl.Columns.Add COLUMN_SENT, "Sent", 150, True
    wndReportControl.Columns.Add COLUMN_SIZE, "Size", 50, True
    wndReportControl.Columns.Add COLUMN_PRICE, "Price", 80, True
    
    wndReportControl.SetImageList ImageListReport

    Dim Record As ReportRecord
    ' Setup ReportControl records
    AddReportRecord Nothing, True, "Undeliverable Mail", "postmaster@mail.codejock.com", "21/06/2004", 7, 5.23
    AddReportRecord Nothing, False, "RE: Hi Mary", "Peter Parker", "19/06/2004", 17, 3.34
    Set Record = AddReportRecord(Nothing, True, "RE:", "Bruce Wayne", "19/06/2004", 11, 3.32)
    AddReportRecord Record, True, "Re: it's me again", "Clark Kent [ckent@codetoolbox.com]", "17/06/2004", 10, 6.34
    AddReportRecord Record, True, "Re: it's me again", "QueryReply", "17/06/2004", 41, 16.42
    Set Record = AddReportRecord(Record, False, "I don't understand:", "Bruce Wayne", "17/06/2004", 1, 5.12)
    AddReportRecord Record, False, "Re:", "Bruce Wayne", "17/06/2004", 23, 8.76
    
    AddReportRecord Nothing, True, "Undeliverable Mail", "postmaster@mail.codejock.com", "21/06/2004", 7, 5.23
    AddReportRecord Nothing, False, "RE: Hi Mary", "Peter Parker", "19/06/2004", 17, 3.34
    Set Record = AddReportRecord(Nothing, True, "RE:", "Bruce Wayne", "19/06/2004", 11, 3.32)
    AddReportRecord Record, False, "Re: it's me again", "Clark Kent [ckent@codetoolbox.com]", "17/06/2004", 10, 6.34
    AddReportRecord Record, True, "Re: it's me again", "QueryReply", "17/06/2004", 41, 16.42
    Set Record = AddReportRecord(Record, True, "I don't understand:", "Bruce Wayne", "17/06/2004", 1, 5.12)
    AddReportRecord Record, True, "Re:", "Bruce Wayne", "17/06/2004", 23, 8.76
    
    AddReportRecord Nothing, False, "Undeliverable Mail", "postmaster@mail.codejock.com", "21/06/2004", 7, 5.23
    AddReportRecord Nothing, True, "RE: Hi Mary", "Peter Parker", "19/06/2004", 17, 3.34
    Set Record = AddReportRecord(Nothing, True, "RE:", "Bruce Wayne", "19/06/2004", 11, 3.32)
    AddReportRecord Record, False, "Re: it's me again", "Clark Kent [ckent@codetoolbox.com]", "17/06/2004", 10, 6.34
    AddReportRecord Record, False, "Re: it's me again", "QueryReply", "17/06/2004", 41, 16.42
    Set Record = AddReportRecord(Record, True, "I don't understand:", "Bruce Wayne", "17/06/2004", 1, 5.12)
    AddReportRecord Record, True, "Re:", "Bruce Wayne", "17/06/2004", 23, 8.76

    ' Apply all those operations
    wndReportControl.Populate
End Sub

Function AddReportRecord(Parent As ReportRecord, Read As Boolean, Subject As String, From As String, Sent As Date, Size As Long, Price As Single) As ReportRecord
 
    ' Adds a record to current ReportControl.
   
    Dim Record As ReportRecord
    
    If Parent Is Nothing Then
        Set Record = wndReportControl.Records.Add()
    Else
        Set Record = Parent.Childs.Add()
    End If
    
    Dim ITEM As ReportRecordItem
    
    Set ITEM = Record.AddItem("")
    ITEM.Icon = IIf(Read, 3, 2)
    
    Set ITEM = Record.AddItem("")
    ITEM.HasCheckbox = True
    ITEM.Checked = False
    
    Record.AddItem Subject
    Record.AddItem From
    Record.AddItem Sent
    Record.AddItem Size
    Set ITEM = Record.AddItem(Price)
    ITEM.Format = "$ %s"
    
    Set AddReportRecord = Record
End Function

Sub FillGrid()
    Dim rstmp  As New ADODB.Recordset
    
    
    
    Set rstmp = Nothing
End Sub

