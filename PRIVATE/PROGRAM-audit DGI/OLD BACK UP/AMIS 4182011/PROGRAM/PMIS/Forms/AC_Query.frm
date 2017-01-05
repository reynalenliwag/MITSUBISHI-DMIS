VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISAC_Query 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Query"
   ClientHeight    =   8760
   ClientLeft      =   2745
   ClientTop       =   2355
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_Query.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14625
   Begin wizButton.cmd cmdPrint 
      Height          =   315
      Left            =   10050
      TabIndex        =   8
      Top             =   90
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      TX              =   "Print"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "AC_Query.frx":030A
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      ScaleHeight     =   405
      ScaleWidth      =   5475
      TabIndex        =   4
      Top             =   60
      Width           =   5475
      Begin VB.OptionButton Option3 
         Caption         =   "By Model Application"
         Height          =   315
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Description"
         Height          =   315
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Product No."
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.TextBox textSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5550
      MaxLength       =   35
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   90
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid grdQUERY2 
      Height          =   4425
      Left            =   60
      TabIndex        =   0
      Top             =   3810
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   7805
      _Version        =   393216
      Cols            =   24
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdDialog 
      Left            =   30
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstParts 
      Height          =   3345
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "AC_Query.frx":0326
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picPartsInquiry 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      ScaleHeight     =   465
      ScaleWidth      =   13785
      TabIndex        =   9
      Top             =   8220
      Width           =   13785
      Begin wizButton.cmd cmdSearchPartNo 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         ToolTipText     =   "Search Accessories by Part Number"
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         TX              =   "F2 - Search Part No."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "AC_Query.frx":0488
      End
      Begin wizButton.cmd cmdPARTSINQUIRYExit 
         Height          =   315
         Left            =   4440
         TabIndex        =   11
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "AC_Query.frx":04A4
      End
      Begin wizButton.cmd cmdTransLedger 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         ToolTipText     =   "View Transaction Ledger"
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         TX              =   "F3 - Trans. Ledger"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "AC_Query.frx":04C0
      End
      Begin wizButton.cmd cmd2 
         Height          =   315
         Left            =   6660
         TabIndex        =   13
         Top             =   60
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         TX              =   "View Un-balance Stocks"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_Query.frx":04DC
      End
      Begin wizButton.cmd cmd3 
         Height          =   315
         Left            =   8580
         TabIndex        =   14
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         TX              =   "Balance Ledger"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   99
         MICON           =   "AC_Query.frx":063E
      End
   End
   Begin VB.Label labAydi 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11550
      TabIndex        =   1
      Top             =   5250
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "frmPMISAC_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                                             As Excel.Application
Dim xlBook                                                            As Excel.Workbook
Dim xlSheet                                                           As Excel.Worksheet
Dim RSPO_HD, rsPO_Hist                                                As ADODB.Recordset
Attribute rsPO_Hist.VB_VarUserMemId = 1073938432
Dim RSPARTMAS                                                         As ADODB.Recordset
Attribute RSPARTMAS.VB_VarUserMemId = 1073938434
Dim rsMatMas, rsMATISS, rsMATISS_HIST, rsMATREC, rsMATREC_HIST        As ADODB.Recordset
Attribute rsMatMas.VB_VarUserMemId = 1073938435
Attribute rsMATISS.VB_VarUserMemId = 1073938435
Attribute rsMATISS_HIST.VB_VarUserMemId = 1073938435
Attribute rsMATREC.VB_VarUserMemId = 1073938435
Attribute rsMATREC_HIST.VB_VarUserMemId = 1073938435
Dim rsRR_HD, rsREC_HIST, rsOrd_Hd                                     As ADODB.Recordset
Attribute rsRR_HD.VB_VarUserMemId = 1073938440
Attribute rsREC_HIST.VB_VarUserMemId = 1073938440
Attribute rsOrd_Hd.VB_VarUserMemId = 1073938440
Dim rsORD_HIST, rsPOSTAT, rsTdayTran                                  As ADODB.Recordset
Attribute rsORD_HIST.VB_VarUserMemId = 1073938443
Attribute rsPOSTAT.VB_VarUserMemId = 1073938443
Attribute rsTdayTran.VB_VarUserMemId = 1073938443
Dim rsDAYTRAN, rsDNPP, rsNEWDNPInc                                    As ADODB.Recordset
Attribute rsDAYTRAN.VB_VarUserMemId = 1073938446
Attribute rsDNPP.VB_VarUserMemId = 1073938446
Attribute rsNEWDNPInc.VB_VarUserMemId = 1073938446
Dim AddorEdit                                                         As String
Attribute AddorEdit.VB_VarUserMemId = 1073938449
Dim kcnt                                                              As Integer
Attribute kcnt.VB_VarUserMemId = 1073938450
Dim gconNewDNPP                                                       As ADODB.Connection
Attribute gconNewDNPP.VB_VarUserMemId = 1073938451

Function PARTSINQUIRYBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                                        As Boolean
    Dim rsBClone                                                      As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = RSPARTMAS.Clone

        rsBClone.Find "STOCKNO = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            RSPARTMAS.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    PARTSINQUIRYBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC from PMIS_Accessories where STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function

Function TRANDETAILSBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                                        As Boolean
    Dim rsBClone                                                      As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsTdayTran.Clone

        rsBClone.Find "tranno = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            rsTdayTran.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    TRANDETAILSBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Sub RiFresh5()
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "Select * from PMIS_TdayTran WHERE [TYPE] = 'A' order by tranno asc ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    FillTRANDETAILSGrid
End Sub

Sub initPARTSINQUIRYGrid()
    If CSMS_PARTSQUERY = True Then
        With lstParts
            .ColumnHeaders(1).Text = "Part #"
            .ColumnHeaders(1).Width = 1500
            .ColumnHeaders(2).Text = "Part Desc"
            .ColumnHeaders(2).Width = 3500
            .ColumnHeaders(3).Text = "Veh.Type"
            .ColumnHeaders(4).Text = "Model Code"
            .ColumnHeaders(5).Text = "On-Hand"
            .ColumnHeaders(5).Width = 1500
            .ColumnHeaders(6).Text = "SRP"
            .ColumnHeaders(6).Width = 1700
            .ColumnHeaders(7).Width = 1
            .ColumnHeaders(8).Width = 1
            .ColumnHeaders(9).Width = 1
            .ColumnHeaders(10).Width = 1
            .ColumnHeaders(11).Width = 1
            .ColumnHeaders(12).Width = 1
            .ColumnHeaders(13).Width = 1
            .ColumnHeaders(14).Width = 1
            .ColumnHeaders(15).Width = 1
            .ColumnHeaders(16).Width = 1
            .ColumnHeaders(17).Width = 1
            .ColumnHeaders(18).Width = 1
            .ColumnHeaders(19).Width = 1
            .ColumnHeaders(20).Width = 1
            .ColumnHeaders(21).Width = 1
            .ColumnHeaders(22).Width = 1
        End With
    Else
        With lstParts
            
            .ColumnHeaders(1).Width = 1500: .ColumnHeaders(1).Text = "Stock #"
            .ColumnHeaders(2).Width = 3000: .ColumnHeaders(2).Text = "Description"
            .ColumnHeaders(3).Width = 1000: .ColumnHeaders(3).Text = "Veh.Type"
            .ColumnHeaders(4).Width = 1000: .ColumnHeaders(4).Text = "Model Code"
            .ColumnHeaders(5).Width = 1000: .ColumnHeaders(5).Text = "Location"
            .ColumnHeaders(6).Width = 800: .ColumnHeaders(6).Text = "Last On-Hand"
            .ColumnHeaders(7).Width = 800: .ColumnHeaders(7).Text = "Last Mac"
            .ColumnHeaders(8).Width = 1000: .ColumnHeaders(8).Text = "On-Hand"
            
            .ColumnHeaders(9).Width = 1000: .ColumnHeaders(9).Text = "Mac"
            .ColumnHeaders(10).Width = 1000: .ColumnHeaders(10).Text = "MAD"
            .ColumnHeaders(11).Width = 1000: .ColumnHeaders(11).Text = "WFP"
            .ColumnHeaders(12).Text = "SRP"
            .ColumnHeaders(13).Width = 1000: .ColumnHeaders(13).Text = "Phy Count"
            .ColumnHeaders(14).Text = "Adj.Count"
            .ColumnHeaders(15).Text = "Total PO"
            .ColumnHeaders(16).Text = "Total RR"
            .ColumnHeaders(17).Text = "Total ISS"
            .ColumnHeaders(18).Text = "MTD RR"
            .ColumnHeaders(19).Text = "MTD ISS"
            .ColumnHeaders(20).Text = "Rank"
            .ColumnHeaders(21).Text = "Years"
            .ColumnHeaders(22).Text = "Check"
        End With
    End If
End Sub

Sub FillPARTSINQUIRYGrid()
    On Error GoTo ERRORCODE
    kcnt = 0
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        lstParts.ZOrder 0
        textSearch.ZOrder 0
        lstParts.Sorted = False: lstParts.Refresh
        lstParts.Enabled = True
        Listview_Loadval lstParts.ListItems, RSPARTMAS
    Else
        lstParts.Enabled = False
        lstParts.Sorted = True: lstParts.Refresh
    End If
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Sub initGrd2()
    Dim KIM                                                           As Integer
    With grdQUERY2
        .Row = 0
        .FormatString = "Tran Type | Tran No. | Tran Date | " & _
                        "Item No. | Part Ordered | Part Supplied | " & _
                        "Quantity | Unit Cost "
        .ColWidth(5) = 2500
        For KIM = 8 To 21
            .ColWidth(KIM) = 1
        Next
    End With
End Sub

Sub initGrd3()
    Dim KIM                                                           As Integer
    With grdQUERY2
        .Row = 0
        .FormatString = "Part Number | Tran Date | Tran No | " & _
                        "Supplier Name/Cust Name | RO Number | Ref. No. | Received | Issued | Balance | " & _
                        "Unit Cost |  MAC  |  EXT. MAC |    SRP | Status | User   "
        .ColWidth(0) = 1300: .ColWidth(2) = 1300: .ColWidth(3) = 3500: .ColWidth(9) = 1000: .ColWidth(10) = 1000: .ColWidth(11) = 1000: .ColWidth(12) = 1000: .ColWidth(13) = 800: .ColWidth(14) = 800
        For KIM = 15 To 23: .ColWidth(KIM) = 1: Next
    End With
End Sub

Sub initGrd4()
    Dim KIM                                                           As Integer
    With grdQUERY2
        .Row = 0
        .FormatString = "Material Code | Tran Date | Tran No | " & _
                        "Supplier Name/Cust Name | Received | Issued | Balance | " & _
                        "Unit Cost |  MAC  | EXT. MAC | Status "
        .ColWidth(0) = 1300: .ColWidth(2) = 1400: .ColWidth(3) = 2500: .ColWidth(7) = 1000: .ColWidth(8) = 1000: .ColWidth(9) = 1000
        For KIM = 11 To 21: .ColWidth(KIM) = 1: Next
    End With
End Sub

Sub FillTRANDETAILSGrid()
    On Error GoTo ERRORCODE
    Dim YzaCnt                                                        As Integer
    YzaCnt = 0
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        Screen.MousePointer = 11
        rsTdayTran.MoveFirst
        Do While Not rsTdayTran.EOF
            YzaCnt = YzaCnt + 1
            rsTdayTran.MoveNext
            DoEvents
        Loop
        Screen.MousePointer = 0
    Else
    End If
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsPartMas2                                                    As ADODB.Recordset
    lstParts.Sorted = False: lstParts.ListItems.Clear
    lstParts.Enabled = False
    Set rsPartMas2 = New ADODB.Recordset
    lstParts.ZOrder 0
    textSearch.ZOrder 0
    If CSMS_PARTSQUERY = True Then
        Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_Accessories  order by STOCKNO asc")
    Else
        Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, location, lastm_oh, lastm_mac, Onhand, Mac, mad, wfp, srp, phycount, adjphycnt, tpoqty, trecqty, tissqty from PMIS_Accessories  order by STOCKNO asc")
    End If
    If Not (rsPartMas2.EOF And rsPartMas2.BOF) Then
        lstParts.Enabled = True
        Listview_Loadval Me.lstParts.ListItems, rsPartMas2
        lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsPartMas2                                                    As ADODB.Recordset
    lstParts.Sorted = False: lstParts.ListItems.Clear
    lstParts.Enabled = False
    Set rsPartMas2 = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    If CSMS_PARTSQUERY = True Then
        If Option1.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_Accessories where STOCKNO like '" & xxx & "%' order by STOCKNO asc")
        End If
        If Option2.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_Accessories where STOCKDESC like '" & xxx & "%' order by STOCKDESC asc")
        End If
        If Option3.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_Accessories where modelcode like '" & xxx & "%' order by modelcode asc")
        End If
    Else
        If Option1.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, location, lastm_oh, lastm_mac, Onhand, Mac, mad, wfp, srp, phycount, adjphycnt, tpoqty, trecqty, tissqty from PMIS_Accessories where STOCKNO like '" & xxx & "%' order by STOCKNO asc")
        End If
        If Option2.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, location, lastm_oh, lastm_mac, Onhand, Mac, mad, wfp, srp, phycount, adjphycnt, tpoqty, trecqty, tissqty from PMIS_Accessories where STOCKDESC like '" & xxx & "%' order by STOCKDESC asc")
        End If
        If Option3.Value = True Then
            Set rsPartMas2 = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, location, lastm_oh, lastm_mac, Onhand, Mac, mad, wfp, srp, phycount, adjphycnt, tpoqty, trecqty, tissqty from PMIS_Accessories where modelcode like '" & xxx & "%' order by modelcode asc")
        End If
    End If
    If Not (rsPartMas2.EOF And rsPartMas2.BOF) Then
        lstParts.Enabled = True
        Listview_Loadval Me.lstParts.ListItems, rsPartMas2
        lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub



Private Sub cmd2_Click()
 cleargrid grdQUERY2
    Dim rsPartMas2                                     As ADODB.Recordset
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsPartMas2 = New ADODB.Recordset
    lstParts.ZOrder 0
    textSearch.ZOrder 0
    grdQUERY2.Visible = True
    Dim SQL                                            As String
    SQL = " SELECT STOCKNO , STOCKDESC, vehtype, modelcode, location, lastm_oh, lastm_mac, Onhand, Mac, mad, wfp, srp, phycount, adjphycnt, tpoqty, trecqty, tissqty from pmis_stockmas WHERE STOCKNO IN "
    SQL = SQL & " (SELECT STOCKNO FROM (SELECT STOCKNO, ONHAND MASTERFILE,"
    SQL = SQL & " ("
    SQL = SQL & " SELECT ISNULL(SUM(TRANQTY),0) FROM ("
    SQL = SQL & " SELECT    TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='BEG' AND TYPE='A' AND STATUS IN('P','B') "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT  (TRANQTY)  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT (TRANQTY)   FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')  "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT (TRANQTY)    FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT (TRANQTY)    FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B') "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & " UNION ALL"
    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & "  ) T) AS LEDGERBALANCE FROM PMIS_STOCKMAS WHERE TYPE='A') C WHERE C.MASTERFILE<>C.LEDGERBALANCE)"

    Set rsPartMas2 = gconDMIS.Execute(SQL)

    If Not (rsPartMas2.EOF And rsPartMas2.BOF) Then
        lstParts.Enabled = True
        Listview_Loadval Me.lstParts.ListItems, rsPartMas2
        lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
    If lstParts.ListItems.Count > 0 Then
        cmd3.Enabled = True
    End If

End Sub

Private Sub cmd3_Click()
 Dim SQL                                            As String

    SQL = " UPDATE PMIS_STOCKMAS SET ONHAND=LEDGERVIEW.LEDGERBALANCE"
    SQL = SQL & " FROM "
    SQL = SQL & " (  SELECT * FROM (SELECT STOCKNO, ONHAND MASTERFILE,("
    SQL = SQL & "    SELECT ISNULL(SUM(TRANQTY),0) FROM ("
    SQL = SQL & "    SELECT    TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='BEG' AND TYPE='A' AND STATUS IN('P','B') "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT  (TRANQTY)  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT (TRANQTY)   FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')  "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT (TRANQTY)    FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT (TRANQTY)    FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B') "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & "    UNION ALL"
    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='A' AND STATUS IN('P','B')  "
    SQL = SQL & "     ) T) AS LEDGERBALANCE FROM PMIS_STOCKMAS WHERE TYPE='A') C WHERE C.MASTERFILE<>C.LEDGERBALANCE"
    SQL = SQL & "    ) LEDGERVIEW"
    SQL = SQL & " INNER JOIN PMIS_STOCKMAS ON LEDGERVIEW.STOCKNO=PMIS_STOCKMAS.STOCKNO"
    gconDMIS.Execute SQL
    FillGrid
    cmd2.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
       
    If grdQUERY2.TextMatrix(1, 0) = "No Entry" Then
        MsgBox "No Record(s) to Print!", vbInformation, "Accessories Stock Cards"
        Screen.MousePointer = 0
        Exit Sub
    End If
       
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\Stock Ledger.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    
    Dim rowCtr, xlrCtr As Long
    xlrCtr = 5
    For rowCtr = 1 To grdQUERY2.Rows - 1
        With grdQUERY2
            xlSheet.Cells(xlrCtr, "A") = .TextMatrix(rowCtr, 0)
            xlSheet.Cells(xlrCtr, "B") = .TextMatrix(rowCtr, 1)
            xlSheet.Cells(xlrCtr, "C") = .TextMatrix(rowCtr, 2)
            xlSheet.Cells(xlrCtr, "D") = .TextMatrix(rowCtr, 3)
            xlSheet.Cells(xlrCtr, "E") = .TextMatrix(rowCtr, 4)
            xlSheet.Cells(xlrCtr, "F") = .TextMatrix(rowCtr, 5)
            xlSheet.Cells(xlrCtr, "G") = .TextMatrix(rowCtr, 6)
            xlSheet.Cells(xlrCtr, "H") = .TextMatrix(rowCtr, 7)
            xlSheet.Cells(xlrCtr, "I") = .TextMatrix(rowCtr, 8)
            xlSheet.Cells(xlrCtr, "J") = .TextMatrix(rowCtr, 9)
            xlSheet.Cells(xlrCtr, "K") = .TextMatrix(rowCtr, 10)
            xlSheet.Cells(xlrCtr, "L") = .TextMatrix(rowCtr, 11)
            xlSheet.Cells(xlrCtr, "M") = .TextMatrix(rowCtr, 12)
            xlSheet.Cells(xlrCtr, "N") = .TextMatrix(rowCtr, 13)
            xlSheet.Cells(xlrCtr, "O") = .TextMatrix(rowCtr, 14)
            xlrCtr = xlrCtr + 1
        End With
    Next
    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0
CloseExcel:
    Set xlApp = Nothing
End Sub

Private Sub cmdTransLedger_Click()
    If lstParts.SelectedItem Is Nothing Then Exit Sub
    
     cmdPrint.Enabled = True
    
    Dim fild, STOCKNUMBER                                             As String
    Dim YzaCnt                                                        As Integer
    Dim rsDAYTRAN2                                                    As ADODB.Recordset
    Dim MovingAverageCost                                             As Double
    YzaCnt = 0
    grdQUERY2.Row = grdQUERY2.Row
    grdQUERY2.Col = 0
    STOCKNUMBER = lstParts.SelectedItem
    grdQUERY2.Col = 17
    fild = grdQUERY2.Text
    grdQUERY2.ZOrder 0
    cleargrid grdQUERY2
    initGrd3
    Dim Balans                                                        As Integer
    If STOCKNUMBER <> "" Then
        Set rsDAYTRAN2 = New ADODB.Recordset
        rsDAYTRAN2.Open "select id,ItemNo,trandate,STOCK_ORD,trantype,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,USERCODE from PMIS_DayTran where TYPE = 'A' AND STOCK_ORD = '" & STOCKNUMBER & "' order by trandate asc, id asc, tranno asc", gconDMIS
        If Not rsDAYTRAN2.EOF And Not rsDAYTRAN2.BOF Then
            rsDAYTRAN2.MoveFirst
            Screen.MousePointer = 11
            Balans = 0
            Do While Not rsDAYTRAN2.EOF
                If Null2String(rsDAYTRAN2!TranType) = "BEG" Or Null2String(rsDAYTRAN2!TranType) = "IN" Then
                    If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                        Balans = Balans + N2Str2IntZero(rsDAYTRAN2!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!trandate) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!TranType) & " #" & Null2String(rsDAYTRAN2!TRANNO) & Chr(9) & _
                                      "BEGINNING" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Balans & Chr(9) & _
                                      N2Str2Zero(rsDAYTRAN2!TRANUCOST) & Chr(9) & _
                                      N2Str2Zero(rsDAYTRAN2!Mac) & Chr(9) & _
                                      Format(N2Str2Zero(rsDAYTRAN2!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsDAYTRAN2!STATUS) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!USERCODE)
                    MovingAverageCost = N2Str2Zero(rsDAYTRAN2!Mac)
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                If Null2String(rsDAYTRAN2!TranType) = "RR" Then
                    Set rsREC_HIST = New ADODB.Recordset
                    rsREC_HIST.Open "select rrno,rrdate,recvd_from,invno from PMIS_REC_Hist where TYPE = 'A' AND rrno = " & N2Str2Null(rsDAYTRAN2!TRANNO), gconDMIS
                    If Not rsREC_HIST.EOF And Not rsREC_HIST.BOF Then
                        If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                            Balans = Balans + N2Str2IntZero(rsDAYTRAN2!tranqty)
                        End If
                        grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                          Null2String(rsREC_HIST!RRDATE) & Chr(9) & _
                                          Null2String(rsDAYTRAN2!TranType) & " #" & Null2String(rsREC_HIST!RRNO) & Chr(9) & _
                                          Null2String(rsREC_HIST!recvd_from) & Chr(9) & _
                                          "" & Chr(9) & _
                                          Null2String(rsREC_HIST!invno) & Chr(9) & _
                                          N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                        0 & Chr(9) & _
                                          Balans & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN2!TRANUCOST) & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN2!Mac) & Chr(9) & _
                                          Format(N2Str2Zero(rsDAYTRAN2!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                        0 & Chr(9) & _
                                          Null2String(rsDAYTRAN2!STATUS) & Chr(9) & _
                                          Null2String(rsDAYTRAN2!USERCODE)
                        MovingAverageCost = N2Str2Zero(rsDAYTRAN2!Mac)
                        YzaCnt = YzaCnt + 1
                        If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                    End If
                End If
                If Null2String(rsDAYTRAN2!TranType) = "OUT" Then
                    If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                        Balans = Balans - N2Str2IntZero(rsDAYTRAN2!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!trandate) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!TranType) & " #" & Null2String(rsDAYTRAN2!TRANNO) & Chr(9) & _
                                      "BEG. OUT" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                    0 & Chr(9) & _
                                      N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                      Balans & Chr(9) & _
                                      Format(0, "0.00") & Chr(9) & _
                                      ToDoubleNumber(MovingAverageCost) & Chr(9) & _
                                      ToDoubleNumber(MovingAverageCost * Balans) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsDAYTRAN2!STATUS) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!USERCODE)
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If

                If Null2String(rsDAYTRAN2!TranType) = "RIV" Or Null2String(rsDAYTRAN2!TranType) = "CSH" Or Null2String(rsDAYTRAN2!TranType) = "CHG" Or Null2String(rsDAYTRAN2!TranType) = "DR" Then
                    Set rsORD_HIST = New ADODB.Recordset
                    rsORD_HIST.Open "select trantype,tranno,trandate,custname,rono from PMIS_Ord_Hist where TYPE='A' AND trantype = " & N2Str2Null(rsDAYTRAN2!TranType) & " AND tranno = " & N2Str2Null(rsDAYTRAN2!TRANNO), gconDMIS
                    If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
                        If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                            Balans = Balans - N2Str2IntZero(rsDAYTRAN2!tranqty)
                        End If
                        grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                          Null2String(rsORD_HIST!trandate) & Chr(9) & _
                                          Null2String(rsORD_HIST!TranType) & " #" & Null2String(rsORD_HIST!TRANNO) & Chr(9) & _
                                          Null2String(rsORD_HIST!custname) & Chr(9) & _
                                          Null2String(rsORD_HIST!RoNo) & Chr(9) & _
                                          "" & Chr(9) & _
                                        0 & Chr(9) & _
                                          N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                          Balans & Chr(9) & _
                                          FormatNumber(NumericVal(rsDAYTRAN2!TRANUCOST)) & Chr(9) & _
                                          FormatNumber(MovingAverageCost) & Chr(9) & _
                                          ToDoubleNumber(MovingAverageCost * Balans) & Chr(9) & _
                                          N2Str2IntZero(rsDAYTRAN2!TRANUPRICE) & Chr(9) & _
                                          Null2String(rsDAYTRAN2!STATUS) & Chr(9) & _
                                          Null2String(rsDAYTRAN2!USERCODE)
                        YzaCnt = YzaCnt + 1
                        If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                    End If
                End If

                If Null2String(rsDAYTRAN2!TranType) = "ADJ" And Null2String(rsDAYTRAN2!IN_OUT) = "O" Then
                    If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                        Balans = Balans - N2Str2IntZero(rsDAYTRAN2!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!trandate) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!TranType) & " #" & Null2String(rsDAYTRAN2!TRANNO) & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                    0 & Chr(9) & _
                                      N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                      Balans & Chr(9) & _
                                      FormatNumber(NumericVal(rsDAYTRAN2!TRANUCOST)) & Chr(9) & _
                                      Format(NumericVal(rsDAYTRAN2!Mac)) & Chr(9) & _
                                      Format(NumericVal(rsDAYTRAN2!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsDAYTRAN2!STATUS)
                                      Null2String (rsDAYTRAN2!USERCODE) & Chr(9) & _
                    MovingAverageCost = FormatNumber(NumericVal(rsDAYTRAN2!Mac))
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                If Null2String(rsDAYTRAN2!TranType) = "ADJ" And Null2String(rsDAYTRAN2!IN_OUT) = "I" Then
                    If Null2String(rsDAYTRAN2!STATUS) <> "C" Then
                        Balans = Balans + N2Str2IntZero(rsDAYTRAN2!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsDAYTRAN2!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!trandate) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!TranType) & " #" & Null2String(rsDAYTRAN2!TRANNO) & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      N2Str2IntZero(rsDAYTRAN2!tranqty) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Balans & Chr(9) & _
                                      FormatNumber(NumericVal(rsDAYTRAN2!TRANUCOST)) & Chr(9) & _
                                      Format(NumericVal(rsDAYTRAN2!Mac)) & Chr(9) & _
                                      Format(NumericVal(rsDAYTRAN2!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsDAYTRAN2!STATUS) & Chr(9) & _
                                      Null2String(rsDAYTRAN2!USERCODE)
                    MovingAverageCost = FormatNumber(NumericVal(rsDAYTRAN2!Mac))
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                DoEvents
                rsDAYTRAN2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Set rsTdayTran = New ADODB.Recordset
        rsTdayTran.Open "select id,ItemNo,STOCK_ORD,trantype,trandate,tranno,itemno,tranqty,tranucost,mac,status,in_out,TRANUPRICE,USERCODE from PMIS_TdayTran where TYPE = 'A' AND STOCK_ORD = '" & STOCKNUMBER & "' order by trandate asc, id asc, tranno asc", gconDMIS
        If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
            rsTdayTran.MoveFirst
            Screen.MousePointer = 11
            Do While Not rsTdayTran.EOF
                If Null2String(rsTdayTran!TranType) = "BEG" Or Null2String(rsTdayTran!TranType) = "IN" Then
                    If Null2String(rsTdayTran!STATUS) <> "C" Then
                        Balans = Balans + N2Str2IntZero(rsTdayTran!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsTdayTran!trandate) & Chr(9) & _
                                      Null2String(rsTdayTran!TranType) & " #" & Null2String(rsTdayTran!TRANNO) & Chr(9) & _
                                      "BEGINNING" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Balans & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!TRANUCOST) & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!Mac) & Chr(9) & _
                                      Format(N2Str2Zero(rsTdayTran!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsTdayTran!STATUS) & Chr(9) & ""
                    MovingAverageCost = N2Str2Zero(rsTdayTran!Mac)
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                If Null2String(rsTdayTran!TranType) = "RR" Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "select rrno,rrdate,recvd_from,invno from PMIS_RR_Hd where TYPE = 'A' AND rrno = " & N2Str2Null(rsTdayTran!TRANNO), gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        If Null2String(rsTdayTran!STATUS) <> "C" Then
                            Balans = Balans + N2Str2IntZero(rsTdayTran!tranqty)
                        End If
                        grdQUERY2.AddItem Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                          Null2String(rsRR_HD!RRDATE) & Chr(9) & _
                                          Null2String(rsTdayTran!TranType) & " #" & Null2String(rsRR_HD!RRNO) & Chr(9) & _
                                          Null2String(rsRR_HD!recvd_from) & Chr(9) & _
                                          "" & Chr(9) & _
                                          Null2String(rsRR_HD!invno) & Chr(9) & _
                                          N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                        0 & Chr(9) & _
                                          Balans & Chr(9) & _
                                          N2Str2Zero(rsTdayTran!TRANUCOST) & Chr(9) & _
                                          N2Str2Zero(rsTdayTran!Mac) & Chr(9) & _
                                          Format(N2Str2Zero(rsTdayTran!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                        0 & Chr(9) & _
                                          Null2String(rsTdayTran!STATUS) & Chr(9) & _
                                          Null2String(rsTdayTran!USERCODE)
                        MovingAverageCost = N2Str2Zero(rsTdayTran!Mac)
                        YzaCnt = YzaCnt + 1
                        If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                    End If
                End If

                If Null2String(rsTdayTran!TranType) = "RIV" Or Null2String(rsTdayTran!TranType) = "CSH" Or Null2String(rsTdayTran!TranType) = "CHG" Or Null2String(rsTdayTran!TranType) = "DR" Or Null2String(rsTdayTran!TranType) = "OUT" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    rsOrd_Hd.Open "select trantype,tranno,trandate,custname,rono from PMIS_Ord_Hd where TYPE = 'A' AND trantype = " & N2Str2Null(rsTdayTran!TranType) & " AND tranno = " & N2Str2Null(rsTdayTran!TRANNO), gconDMIS
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        If Null2String(rsTdayTran!STATUS) <> "C" Then
                            Balans = Balans - N2Str2IntZero(rsTdayTran!tranqty)
                        End If
                        grdQUERY2.AddItem Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                          Null2String(rsOrd_Hd!trandate) & Chr(9) & _
                                          Null2String(rsOrd_Hd!TranType) & " #" & Null2String(rsOrd_Hd!TRANNO) & Chr(9) & _
                                          Null2String(rsOrd_Hd!custname) & Chr(9) & _
                                          "" & Chr(9) & _
                                          Null2String(rsOrd_Hd!RoNo) & Chr(9) & _
                                        0 & Chr(9) & _
                                          N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                          Balans & Chr(9) & _
                                          FormatNumber(NumericVal(rsTdayTran!TRANUCOST)) & Chr(9) & _
                                          ToDoubleNumber(MovingAverageCost) & Chr(9) & _
                                          ToDoubleNumber(MovingAverageCost * Balans) & Chr(9) & _
                                          N2Str2IntZero(rsTdayTran!TRANUPRICE) & Chr(9) & _
                                          Null2String(rsTdayTran!STATUS) & Chr(9) & _
                                          Null2String(rsTdayTran!USERCODE)
                        YzaCnt = YzaCnt + 1
                        If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                    End If
                End If

                If Null2String(rsTdayTran!TranType) = "ADJ" And Null2String(rsTdayTran!IN_OUT) = "O" Then
                    If Null2String(rsTdayTran!STATUS) <> "C" Then
                        Balans = Balans - N2Str2IntZero(rsTdayTran!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsTdayTran!trandate) & Chr(9) & _
                                      Null2String(rsTdayTran!TranType) & " #" & Null2String(rsTdayTran!TRANNO) & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                    0 & Chr(9) & _
                                      N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                      Balans & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!TRANUCOST) & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!Mac) & Chr(9) & _
                                      Format(N2Str2Zero(rsTdayTran!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsTdayTran!STATUS) & Chr(9) & _
                                      Null2String(rsTdayTran!USERCODE)
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                If Null2String(rsTdayTran!TranType) = "ADJ" And Null2String(rsTdayTran!IN_OUT) = "I" Then
                    If Null2String(rsTdayTran!STATUS) <> "C" Then
                        Balans = Balans + N2Str2IntZero(rsTdayTran!tranqty)
                    End If
                    grdQUERY2.AddItem Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                      Null2String(rsTdayTran!trandate) & Chr(9) & _
                                      Null2String(rsTdayTran!TranType) & " #" & Null2String(rsTdayTran!TRANNO) & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Balans & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!TRANUCOST) & Chr(9) & _
                                      N2Str2Zero(rsTdayTran!Mac) & Chr(9) & _
                                      Format(N2Str2Zero(rsTdayTran!Mac) * Balans, MAXIMUM_DIGIT) & Chr(9) & _
                                    0 & Chr(9) & _
                                      Null2String(rsTdayTran!STATUS) & Chr(9) & _
                                      Null2String(rsTdayTran!USERCODE)
                    MovingAverageCost = N2Str2Zero(rsTdayTran!Mac)
                    YzaCnt = YzaCnt + 1
                    If YzaCnt = 1 Then grdQUERY2.RemoveItem 1
                End If
                DoEvents
                rsTdayTran.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        If YzaCnt > 6 Then grdQUERY2.TopRow = YzaCnt - 5
        
        Call NEW_LogAudit("I", "ACCESSORIES LEDGER FILE", "", "", "Accessories", "ACC NO: " & STOCKNUMBER, "", "")
    Else
        MsgSpeechBox "No Transaction on Selected Accessories..."
        Exit Sub
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCESSORIES LEDGER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACCESSORIES LEDGER FILE", "PRINTING")
            
        Case vbKeyEscape
            grdQUERY2.ZOrder 1
            cmdPrint.Enabled = False
        Case vbKeyF2
            cmdSearchPARTNO_Click
        Case vbKeyF3
            cmdTransLedger_Click
        Case vbKeyF4
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    picPartsInquiry.Visible = False
    grdQUERY2.ZOrder 1
    textSearch.Text = ""
    Me.Caption = "ACCESSORIES INVENTORY LEDGER"
    FillGrid
    picPartsInquiry.Visible = True
    picPartsInquiry.ZOrder 0
    initPARTSINQUIRYGrid
End Sub

Private Sub cmdPARTSINQUIRYExit_Click()
    Unload Me
End Sub

Private Sub cmdSearchPARTNO_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub




Private Sub lstParts_DblClick()
cmdTransLedger_Click
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub Option2_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub Option3_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) <> "" Then
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstParts.ListItems.Count > 0 And lstParts.Enabled = True Then: lstParts.SetFocus
    End If
End Sub

