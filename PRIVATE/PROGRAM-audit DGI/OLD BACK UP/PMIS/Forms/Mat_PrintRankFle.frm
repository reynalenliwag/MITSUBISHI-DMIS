VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmCSMSMat_PrintRankfle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Inventory Ranking"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Mat_PrintRankFle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   3660
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
      Height          =   750
      Left            =   1350
      MouseIcon       =   "Mat_PrintRankFle.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Mat_PrintRankFle.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   735
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
      Height          =   750
      Left            =   2115
      MouseIcon       =   "Mat_PrintRankFle.frx":1433
      MousePointer    =   99  'Custom
      Picture         =   "Mat_PrintRankFle.frx":1585
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   825
      Width           =   735
   End
   Begin VB.ComboBox cboDate_Gen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select date of report "
      Top             =   90
      Width           =   2205
   End
   Begin SHDocVwCtl.WebBrowser browRank 
      Height          =   3945
      Left            =   60
      TabIndex        =   2
      Top             =   2310
      Width           =   8685
      ExtentX         =   15319
      ExtentY         =   6959
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "Summary Only"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   930
      TabIndex        =   1
      Top             =   510
      Width           =   2205
   End
   Begin Crystal.CrystalReport rptPrintRankfle 
      Left            =   3090
      Top             =   810
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AS OF:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmCSMSMat_PrintRankfle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRANKFLE            As ADODB.Recordset
Dim rsRankfle2           As ADODB.Recordset
Dim rsProfile            As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If IsDate(cboDate_Gen.Text) = True Then
        Set rsRANKFLE = New ADODB.Recordset
        rsRANKFLE.Open "select * from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
            Screen.MousePointer = 11
            If chkSummary.Value = 1 Then
                'rptPrintRankfle.ReportTitle = "SUMMARY OF INVENTORY RANKING REPORT"
                'PrintSQLReport rptPrintRankfle, CSMS_REPORT_PATH & "ranksum.rpt", "", DMIS_REPORT_Connection, 1
                RANKSUMPRINTING
            Else
                rptPrintRankfle.ReportTitle = "MATERIALS INVENTORY RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

                PrintSQLReport rptPrintRankfle, CSMS_REPORT_PATH & "Mat_ranking.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            End If
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "Not Yet Generated!"
        End If
    Else
        MsgSpeechBox "Invalid Date Generated!"
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Set rsRANKFLE = New ADODB.Recordset
    rsRANKFLE.Open "select date_gen from PMIS_RankFle group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
        cboDate_Gen.Clear
        Do While Not rsRANKFLE.EOF
            cboDate_Gen.AddItem Null2Date(rsRANKFLE!date_gen)
            rsRANKFLE.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Sub RANKSUMPRINTING()
    Screen.MousePointer = 11
    Dim RANKA1QTY, RANKA2QTY, RANKA3QTY, RANKBQTY, RANKCQTY, RANKDQTY, RANKE1QTY, RANKE2QTY, RANKE3QTY, RANKE4QTY, RANKE5QTY, RANKFQTY As Integer
    Dim RANKA1EXTCOST, RANKA2EXTCOST, RANKA3EXTCOST, RANKBEXTCOST, RANKCEXTCOST, RANKDEXTCOST, RANKE1EXTCOST, RANKE2EXTCOST, RANKE3EXTCOST, RANKE4EXTCOST, RANKE5EXTCOST, RANKFEXTCOST As Double
    Dim INVENTORYCLASS, INVENTORYSUBCLASS As String
    Dim INVENTORYONHAND  As Integer
    Dim INVENTORYCOST    As Double
    Dim TOTALINVENTORYONHAND As Integer
    Dim TOTALINVENTORYCOST As Double

    INVENTORYONHAND = 0: INVENTORYCOST = 0
    TOTALINVENTORYONHAND = 0: TOTALINVENTORYCOST = 0
    RANKA1QTY = 0: RANKA2QTY = 0: RANKA3QTY = 0: RANKBQTY = 0: RANKCQTY = 0: RANKDQTY = 0: RANKE1QTY = 0: RANKE2QTY = 0: RANKE3QTY = 0: RANKE4QTY = 0: RANKE5QTY = 0: RANKFQTY = 0
    RANKA1EXTCOST = 0: RANKA2EXTCOST = 0: RANKA3EXTCOST = 0: RANKBEXTCOST = 0: RANKCEXTCOST = 0: RANKDEXTCOST = 0: RANKE1EXTCOST = 0: RANKE2EXTCOST = 0: RANKE3EXTCOST = 0: RANKE4EXTCOST = 0: RANKE5EXTCOST = 0: RANKFEXTCOST = 0

    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_Profile", gconDMIS
    Set rsRANKFLE = New ADODB.Recordset
    rsRANKFLE.Open "select month_gen,onhand,COST,invclass,subinvclas,date_gen from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "' order by invclass,subinvclas asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
        Open CSMS_REPORT_PATH & "RANKSUM.HTML" For Output As #1
        Print #1, "<html><body>"
        Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "<td align=center width=60%>&nbsp;</td>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
        Print #1, "<td align=center width=60%><font size=4 FACE=TIMES NEW ROMAN>" & rsProfile!CompanyName & "</font></td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=15%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
        Print #1, "<td align=center width=70%><font size=4 FACE=TIMES NEW ROMAN><strong>MATERIALS SUMMARY OF INVENTORY RANKING REPORT</strong></font></td>"
        Print #1, "<td align=left width=15%>&nbsp;</td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "<td align=center width=60%>&nbsp;</td>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<FONT SIZE=3 FACE=TIMES NEW ROMAN<b>FOR THE MONTH OF " & The_month(N2Str2IntZero(rsRANKFLE!Month_Gen)) & " " & Year(LOGDATE) & "</b></FONT><br>"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------<br>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left>&nbsp;</td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>QUANTITY</b></FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>TOTAL COST</b></FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>% QTY</b></FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>% COST</b></FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------<br>"
        rsRANKFLE.MoveFirst
        Do While Not rsRANKFLE.EOF
            INVENTORYCLASS = Null2String(rsRANKFLE!InvClass)
            INVENTORYSUBCLASS = Null2String(rsRANKFLE!SubInvClas)
            INVENTORYONHAND = N2Str2IntZero(rsRANKFLE!ONHAND)
            INVENTORYCOST = N2Str2Zero(rsRANKFLE!COST)
            TOTALINVENTORYONHAND = TOTALINVENTORYONHAND + N2Str2IntZero(rsRANKFLE!ONHAND)
            TOTALINVENTORYCOST = TOTALINVENTORYCOST + (INVENTORYONHAND * INVENTORYCOST)
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "1" Then
                RANKA1QTY = RANKA1QTY + INVENTORYONHAND
                RANKA1EXTCOST = RANKA1EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "2" Then
                RANKA2QTY = RANKA2QTY + INVENTORYONHAND
                RANKA2EXTCOST = RANKA2EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "3" Then
                RANKA3QTY = RANKA3QTY + INVENTORYONHAND
                RANKA3EXTCOST = RANKA3EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "B" And INVENTORYSUBCLASS = "" Then
                RANKBQTY = RANKBQTY + INVENTORYONHAND
                RANKBEXTCOST = RANKBEXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "C" And INVENTORYSUBCLASS = "" Then
                RANKCQTY = RANKCQTY + INVENTORYONHAND
                RANKCEXTCOST = RANKCEXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "D" And INVENTORYSUBCLASS = "" Then
                RANKDQTY = RANKDQTY + INVENTORYONHAND
                RANKDEXTCOST = RANKDEXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "1" Then
                RANKE1QTY = RANKE1QTY + INVENTORYONHAND
                RANKE1EXTCOST = RANKE1EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "2" Then
                RANKE2QTY = RANKE2QTY + INVENTORYONHAND
                RANKE2EXTCOST = RANKE2EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "3" Then
                RANKE3QTY = RANKE3QTY + INVENTORYONHAND
                RANKE3EXTCOST = RANKE3EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "4" Then
                RANKE4QTY = RANKE4QTY + INVENTORYONHAND
                RANKE4EXTCOST = RANKE4EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "5" Then
                RANKE5QTY = RANKE5QTY + INVENTORYONHAND
                RANKE5EXTCOST = RANKE5EXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "F" And INVENTORYSUBCLASS = "" Then
                RANKFQTY = RANKFQTY + INVENTORYONHAND
                RANKFEXTCOST = RANKFEXTCOST + (INVENTORYCOST * INVENTORYONHAND)
            End If
            rsRANKFLE.MoveNext
        Loop
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A1 - SUPER FAST MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA1QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA1EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA1QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA1EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A2 - SUPER FAST MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA2QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA2EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA2QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA2EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A3 - SUPER FAST MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA3QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA3EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA3QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA3EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK B &nbsp;&nbsp;- FAST MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKBQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKBEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKBQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKBEXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK C &nbsp;&nbsp;- MEDIUM MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKCQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKCEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKCQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKCEXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK D &nbsp;&nbsp;- SLOW MOVING MATERIALS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKDQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKDEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKDQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKDEXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E1 - NON MOVING MATERIALS FOR 1 YEAR</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE1QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE1EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE1QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE1EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E2 - NON MOVING MATERIALS FOR 2 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE2QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE2EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE2QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE2EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E3 - NON MOVING MATERIALS FOR 3 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE3QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE3EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE3QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE3EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E4 - NON MOVING MATERIALS FOR 4 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE4QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE4EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE4QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE4EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E5 - NON MOVING MATERIALS FOR 5 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE5QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE5EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE5QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE5EXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK F &nbsp;&nbsp;- NEW ITEMS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKFQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKFEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKFQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKFEXTCOST / TOTALINVENTORYCOST) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------<br>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>GRAND TOTAL</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(TOTALINVENTORYONHAND, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(TOTALINVENTORYCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>100.00 %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>100.00 %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open CSMS_REPORT_PATH & "RANKSUM.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRank.Navigate CSMS_REPORT_PATH & "RANKSUM.HTML"
            DoEvents
            browRank.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set rsProfile = Nothing
    Screen.MousePointer = 0
End Sub
