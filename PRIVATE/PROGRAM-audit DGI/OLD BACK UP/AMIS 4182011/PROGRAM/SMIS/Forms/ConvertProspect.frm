VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_MIS_ConvertProspect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Prospect"
   ClientHeight    =   7005
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "ConvertProspect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   9465
   Begin VB.PictureBox picProspect 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   30
      ScaleHeight     =   6615
      ScaleWidth      =   9540
      TabIndex        =   0
      Top             =   390
      Width           =   9540
      Begin XtremeReportControl.ReportControl lstProspect 
         Height          =   5655
         Left            =   90
         TabIndex        =   1
         Top             =   795
         Width           =   9315
         _Version        =   655364
         _ExtentX        =   16431
         _ExtentY        =   9975
         _StockProps     =   64
         BorderStyle     =   2
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6270
         MouseIcon       =   "ConvertProspect.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "ConvertProspect.frx":0A1C
         TabIndex        =   19
         Top             =   300
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   3495
         TabIndex        =   3
         Top             =   322
         Width           =   2760
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   75
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   345
         Width           =   3285
      End
      Begin VB.Label Label3 
         Caption         =   "Search Keyword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3510
         TabIndex        =   20
         Top             =   30
         Width           =   3195
      End
      Begin VB.Label Label2 
         Caption         =   "Select Prospect"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   75
         Width           =   3195
      End
   End
   Begin VB.PictureBox picCus 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   30
      ScaleHeight     =   6615
      ScaleWidth      =   9540
      TabIndex        =   5
      Top             =   390
      Width           =   9540
      Begin VB.OptionButton Option1 
         Caption         =   "Into Existing Customer"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   90
         Width           =   2235
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Into New Customer"
         Height          =   255
         Left            =   2370
         TabIndex        =   12
         Top             =   90
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8490
         MouseIcon       =   "ConvertProspect.frx":0D74
         MousePointer    =   99  'Custom
         Picture         =   "ConvertProspect.frx":0EC6
         TabIndex        =   11
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         MouseIcon       =   "ConvertProspect.frx":121E
         MousePointer    =   99  'Custom
         Picture         =   "ConvertProspect.frx":1370
         TabIndex        =   10
         Top             =   0
         Width           =   945
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5685
         Left            =   30
         ScaleHeight     =   5685
         ScaleWidth      =   9420
         TabIndex        =   7
         Top             =   930
         Width           =   9420
         Begin XtremeReportControl.ReportControl lstCust 
            Height          =   4995
            Left            =   60
            TabIndex        =   8
            Top             =   450
            Width           =   9270
            _Version        =   655364
            _ExtentX        =   16351
            _ExtentY        =   8811
            _StockProps     =   64
            BorderStyle     =   2
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2940
            TabIndex        =   9
            Top             =   30
            Width           =   3285
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Type Any KeyWord To Search"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            TabIndex        =   18
            Top             =   30
            Width           =   2865
         End
      End
      Begin VB.Label LabAcctname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   6990
         TabIndex        =   17
         Top             =   570
         Width           =   2445
      End
      Begin VB.Label labMobile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2010
         TabIndex        =   16
         Top             =   570
         Width           =   2445
      End
      Begin VB.Label labEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4470
         TabIndex        =   15
         Top             =   570
         Width           =   2505
      End
      Begin VB.Label labPhoneno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   14
         Top             =   570
         Width           =   1965
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   390
      Left            =   -150
      TabIndex        =   6
      Top             =   0
      Width           =   9765
      _Version        =   655364
      _ExtentX        =   17224
      _ExtentY        =   688
      _StockProps     =   14
      Caption         =   "::Convert Prospects Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmSMIS_MIS_ConvertProspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prosID                                                            As Long
Dim CusID                                                             As String

Sub FillProspect(XXX As String)
    Dim SQL                                                           As String

    SQL = "select Convert(varchar,  loginitialinquiry , 101) ,   "
    SQL = SQL & "ACCTNAME ,  "
    SQL = SQL & "TELEPHONE + isnull('/'+MOBILE,''),  "
    SQL = SQL & "Variant ,  "
    SQL = SQL & "SAE ,  "
    SQL = SQL & "ProspectID   "
    SQL = SQL & "from CRIS_PROSPECTS WHERE STATUS='O' and isnull(LogSO,0)=0 and CUSCDE is null"
    If XXX <> vbNullString Then
        SQL = SQL & " AND " & XXX
    End If

    flex_FillReportView gconDMIS.Execute(SQL), lstProspect

End Sub

Sub setReportControl(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        '.PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots  'xtpGridSmallDots    '
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        '.PaintManager.GroupForeColor = vbBlue
        .PaintManager.TextFont.Name = "Aril"
        .PaintManager.TextFont.Size = 9
        '.PaintManager.ColumnStyle = xtpColumnOffice2003
    End With


End Sub

Private Sub cmdNext_Click()
    prosID = lstProspect.SelectedRows(0).Record(5).Value
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select * from cris_prospects where prospectid =" & prosID)

    If Not TEMPRS.BOF Or Not TEMPRS.EOF Then
        labPhoneno = TEMPRS!Telephone & ""
        labEmail = TEMPRS!EMAIL & ""
        labMobile = TEMPRS!Mobile & ""
        LABACCTNAME = TEMPRS!AcctName & ""
    End If


    picProspect.Visible = False
    picCus.Visible = True

    'Dim temprs As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select  cuscde, acctname , address, email, phone, mobile from  CRIS_vw_AllProfile order by 2")
    flex_FillReportView TEMPRS, lstCust
    If lstCust.Records.Count > 0 Then
        Option1.Enabled = True
        Option1.Value = False
    Else
        Option1.Enabled = False
    End If

    Text2 = LABACCTNAME


End Sub

Private Sub cmdPrevious_Click()
    picCus.Visible = False
    picProspect.Visible = True
End Sub

Private Sub Command2_Click()
    picCus.Visible = False: picProspect.Visible = True
End Sub

Private Sub Command3_Click()
    If Option2.Value = True Then
        If MsgBox(" Do You Want Convert This Prospect Into New Customer", vbYesNo) = vbYes Then
            frmAllCustomer.AddCustomerFromProspect gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & prosID), ""
            frmAllCustomer.Show
        End If
    Else
        If MsgBox(" Do You Want Convert This Prospect Into Selected  Customer", vbYesNo) = vbYes Then

        End If
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    picProspect.Visible = True: picCus.Visible = False
    ReportControlAddColumnHeader lstProspect, "Date, ProspectName, Phone, Model, SAE"
    ReportControlAddColumnHeader lstCust, "Account Code, Account Name , Address, Telephone,TEL2,FLD3,FLD4"

    ResizeColumnHeader lstCust, "10 , 30, 30, 20,10,10"
    setReportControl lstProspect
    FillProspect ""
    With Combo1
        .AddItem "ACCTNAME"
        .AddItem "VARIANT"
        .AddItem "SAE"
        .AddItem "TELEPHONE"
        .AddItem "MOBILE"
        .ListIndex = 0
    End With
End Sub

Private Sub TabControl1_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
    'checkok here
    '    Cancel = True
End Sub

Private Sub LabAcctname_Click()
    Text2 = LABACCTNAME
End Sub

Private Sub labEmail_Click()
    Text2 = labEmail
End Sub

Private Sub labMobile_Click()
    Text2 = labMobile
End Sub

Private Sub labPhoneno_Click()

    Text2 = labPhoneno
End Sub

Private Sub Text1_Change()
    If Combo1.ListIndex = -1 Then Exit Sub
    FillProspect Combo1.Text & " Like '%" & Replace(Text1, "'", "") & "%'"
End Sub

Private Sub Text2_Change()
    lstCust.FilterText = Text2
    lstCust.Populate
    Exit Sub
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select top 100 cuscde, acctname , address, email, phone, mobile from  CRIS_vw_AllProfile WHERE acctname like '%" & Text2 & "%' order by 2")
    flex_FillReportView TEMPRS, lstCust
    If lstCust.Records.Count > 0 Then
        Option1.Enabled = True
        Option1.Value = False
    Else
        Option1.Enabled = False
    End If
End Sub

