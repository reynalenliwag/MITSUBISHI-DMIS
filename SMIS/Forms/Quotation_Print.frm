VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSMIS_Trans_Quotation_Print 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Quotation"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "Quotation_Print.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   5880
   Begin Crystal.CrystalReport rptPrintRankfle 
      Left            =   4830
      Top             =   5100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton Option2 
         Caption         =   "With out Amortization Details"
         Height          =   285
         Left            =   960
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5010
         Width           =   2835
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Amortization Details"
         Height          =   285
         Left            =   960
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4770
         Value           =   -1  'True
         Width           =   2505
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3060
         MouseIcon       =   "Quotation_Print.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "Quotation_Print.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Close Window"
         Top             =   5430
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Financing Option"
         Height          =   225
         Left            =   420
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4530
         Width           =   2205
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Cash Option"
         Height          =   225
         Left            =   450
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5310
         Width           =   1785
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Vehicles Details From Inventory List"
         Height          =   375
         Left            =   2130
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Add Vehicle Details From Inventory List"
         Top             =   1440
         Width           =   3645
      End
      Begin RichTextLib.RichTextBox rtfText1 
         Height          =   1185
         Left            =   60
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   2090
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Quotation_Print.frx":13DF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfText2 
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   2143
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Quotation_Print.frx":1456
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfText3 
         Height          =   1185
         Left            =   60
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1830
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   2090
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Quotation_Print.frx":14CD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2340
         MouseIcon       =   "Quotation_Print.frx":1544
         MousePointer    =   99  'Custom
         Picture         =   "Quotation_Print.frx":1696
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Print this Record"
         Top             =   5430
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Quotation Footer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Vehicles Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   1530
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Quotation Header"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1965
      End
   End
   Begin VB.PictureBox picViewVehicles 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6285
      Left            =   0
      ScaleHeight     =   6285
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   4680
         Left            =   45
         TabIndex        =   1
         Top             =   810
         Width           =   5670
         _Version        =   655364
         _ExtentX        =   10001
         _ExtentY        =   8255
         _StockProps     =   64
         BorderStyle     =   4
         SkipGroupsFocus =   0   'False
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   0
         Left            =   4980
         MouseIcon       =   "Quotation_Print.frx":1B35
         MousePointer    =   99  'Custom
         Picture         =   "Quotation_Print.frx":1C87
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   5520
         Width           =   705
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4290
         MouseIcon       =   "Quotation_Print.frx":1FC5
         MousePointer    =   99  'Custom
         Picture         =   "Quotation_Print.frx":2117
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Select"
         Top             =   5520
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5550
         TabIndex        =   3
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   1770
         TabIndex        =   2
         Top             =   420
         Width           =   3915
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   -15
         TabIndex        =   19
         Top             =   0
         Width           =   5925
         _Version        =   655364
         _ExtentX        =   10451
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Preview Vehicles"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   135
         TabIndex        =   4
         Top             =   450
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_Quotation_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ENTRY_LOGID                                                       As Long
Dim CustName, CustAdd, CustContact
Attribute CustAdd.VB_VarUserMemId = 1073938433
Attribute CustContact.VB_VarUserMemId = 1073938433
Sub PrintQuotation(xxID As Long, xCustName, xCustAdd, xCustContact)
    ENTRY_LOGID = xxID
    CustName = xCustName
    CustAdd = xCustAdd
    CustContact = xCustContact
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        'PrintAmort = "Y"
        Option1.Enabled = True: Option2.Enabled = True


    Else
        Option1.Enabled = False: Option2.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picViewVehicles, False, Picture1
End Sub

'Upating Code       : AXP-0707200713:28
Private Sub cmdPrint_Click()
    
    On Error GoTo Errorcode:

    If Check1.Value = 0 And Check2.Value = 0 Then
        ShowIsRequiredMsg "Please Select At Least One Option / Check Box"

        Exit Sub
    End If
    Dim filter                                                        As String
    gconDMIS.Execute "UPDATE CRIS_QuotationDocument SET NOTESTEXT=" & N2Str2Null(rtfText3.Text) & " , HEADERTEXT=" & N2Str2Null(rtfText1.Text) & " , FOOTERTEXT=" & N2Str2Null(rtfText2.Text)
    With rptPrintRankfle
        .Formulas(0) = "Sal1 = '" & CustName & "'"
        .Formulas(1) = "Sal2 = '" & CustAdd & "'"
        .Formulas(2) = "Sal3 = '" & CustContact & "'"
        .WindowTitle = " Quotation"



    End With

    If Check1.Value = 0 And Check2.Value = 1 Then
        PrintSQLReport rptPrintRankfle, SMIS_REPORT_PATH & "QuotationCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
    ElseIf Check1.Value = 1 And Check2.Value = 0 Then
        If Option1.Value = True Then
            rptPrintRankfle.Formulas(3) = "PrintAmort = 'Y'"
        Else
            rptPrintRankfle.Formulas(3) = "PrintAmort = 'N'"
        End If
        PrintSQLReport rptPrintRankfle, SMIS_REPORT_PATH & "QuotationFin.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
        If Option1.Value = True Then
            rptPrintRankfle.Formulas(3) = "PrintAmort = 'Y'"
        Else
            rptPrintRankfle.Formulas(3) = "PrintAmort = 'N'"
        End If
        PrintSQLReport rptPrintRankfle, SMIS_REPORT_PATH & "QuotationFinCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
    End If





    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSelectViewVehicles_Click()

    ShowHidePictureBox2 picViewVehicles, False, Picture1
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim myVal                                                         As String

    Set TEMPRS = gconDMIS.Execute("SELECT CODE, DESCRIPTION ,ISFREE FROM SMIS_MRRINV_DETAIL where IGNKEYNO='" & lvViewVehicles.SelectedRows.Row(0).Record(1).Value & "' ORDER BY ISFREE ASC")

    If (TEMPRS.EOF Or TEMPRS.BOF) Then
        Exit Sub
    End If

    While Not TEMPRS.EOF
        If IsNull(TEMPRS.Fields("DESCRIPTION").Value) = False Then
            If TEMPRS.Fields("ISFREE") = True Then
                myVal = myVal & TEMPRS!Description & "(*)" & " , "
            Else
                myVal = myVal & TEMPRS!Description & " , "
            End If

        End If
        TEMPRS.MoveNext
    Wend
    rtfText3 = Left(myVal, Len(myVal) - 3)

End Sub

Private Sub Command1_Click()
    ReportControlAddColumnHeader lvViewVehicles, "DESCRIPTION, CS#, VINO"
    ReportControlPaintManager lvViewVehicles
    flex_FillReportView gconDMIS.Execute("select  Descript, ignkey,  Vino,color , ID from SMIS_MRRINV ORDER BY MODEL "), lvViewVehicles
    ShowHidePictureBox2 picViewVehicles, True, Picture1
    On Error Resume Next
    txtFilterViewVehicles.SetFocus
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT * FROM CRIS_QUOTATIONDOCUMENT")

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        rtfText1.Text = TEMPRS(1)
        rtfText2.Text = TEMPRS(2)
        rtfText3.Text = TEMPRS(3)
    End If
End Sub

Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSelectViewVehicles_Click
    End If
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then Exit Sub
    cmdSelectViewVehicles_Click
End Sub

