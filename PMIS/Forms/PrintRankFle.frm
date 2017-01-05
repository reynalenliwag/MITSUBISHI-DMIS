VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_PrintRankfle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Ranking Report"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "PrintRankFle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   4425
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -360
      TabIndex        =   12
      Top             =   2430
      Width           =   5355
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Inventory Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   5
      Top             =   2070
      Width           =   3795
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Gross Profit Margin Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   1710
      Width           =   3795
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Cost Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   3
      Top             =   1320
      Width           =   3795
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Shelf Stock Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   2
      Top             =   930
      Width           =   3795
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Piece Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   540
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sales Ranking Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Value           =   -1  'True
      Width           =   3795
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
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Choose date from list"
      Top             =   2820
      Width           =   2205
   End
   Begin SHDocVwCtl.WebBrowser browRank 
      Height          =   3945
      Left            =   6180
      TabIndex        =   10
      Top             =   7050
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
      Left            =   1260
      TabIndex        =   7
      Top             =   3270
      Width           =   2205
   End
   Begin Crystal.CrystalReport rptPrintRankfle 
      Left            =   3180
      Top             =   3840
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
      Height          =   795
      Left            =   1950
      MouseIcon       =   "PrintRankFle.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "PrintRankFle.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close Window"
      Top             =   3660
      Width           =   645
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
      Height          =   795
      Left            =   1320
      MouseIcon       =   "PrintRankFle.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "PrintRankFle.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print Report"
      Top             =   3660
      Width           =   645
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
      Left            =   630
      TabIndex        =   11
      Top             =   2850
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_PrintRankfle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRANKFLE                                          As ADODB.Recordset
Dim rsSTKSTAT                                          As ADODB.Recordset
Dim RSPROFILE                                          As ADODB.Recordset

Sub RANKSUMPRINTING()
    Screen.MousePointer = 11
    Dim RANKA1QTY, RANKA2QTY, RANKA3QTY, RANKBQTY, RANKCQTY, RANKDQTY, RANKE1QTY, RANKE2QTY, RANKE3QTY, RANKE4QTY, RANKE5QTY, RANKFQTY As Integer
    Dim RANKA1EXTCOST, RANKA2EXTCOST, RANKA3EXTCOST, RANKBEXTCOST, RANKCEXTCOST, RANKDEXTCOST, RANKE1EXTCOST, RANKE2EXTCOST, RANKE3EXTCOST, RANKE4EXTCOST, RANKE5EXTCOST, RANKFEXTCOST As Double
    Dim INVENTORYCLASS, INVENTORYSUBCLASS              As String
    Dim INVENTORYONHAND                                As Integer
    Dim INVENTORYMAC                                   As Double
    Dim TOTALINVENTORYONHAND                           As Integer
    Dim TOTALINVENTORYMAC                              As Double

    INVENTORYONHAND = 0: INVENTORYMAC = 0
    TOTALINVENTORYONHAND = 0: TOTALINVENTORYMAC = 0
    RANKA1QTY = 0: RANKA2QTY = 0: RANKA3QTY = 0: RANKBQTY = 0: RANKCQTY = 0: RANKDQTY = 0: RANKE1QTY = 0: RANKE2QTY = 0: RANKE3QTY = 0: RANKE4QTY = 0: RANKE5QTY = 0: RANKFQTY = 0
    RANKA1EXTCOST = 0: RANKA2EXTCOST = 0: RANKA3EXTCOST = 0: RANKBEXTCOST = 0: RANKCEXTCOST = 0: RANKDEXTCOST = 0: RANKE1EXTCOST = 0: RANKE2EXTCOST = 0: RANKE3EXTCOST = 0: RANKE4EXTCOST = 0: RANKE5EXTCOST = 0: RANKFEXTCOST = 0

    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile", gconDMIS
    Set rsRANKFLE = New ADODB.Recordset
    rsRANKFLE.Open "select month_gen,onhand,mac,invclass,subinvclas,date_gen from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "' and [TYPE] = 'P' order by invclass,subinvclas asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
        Open PMIS_REPORT_PATH & "RANKSUM.HTML" For Output As #1
        Print #1, "<html><body>"
        Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "<td align=center width=60%>&nbsp;</td>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
        Print #1, "<td align=center width=60%><font size=4 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
        Print #1, "</tr>"
        Print #1, "<tr>"
        Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
        Print #1, "<td align=center width=60%><font size=4 FACE=TIMES NEW ROMAN><strong>SUMMARY OF INVENTORY RANKING REPORT</strong></font></td>"
        Print #1, "<td align=left width=20%>&nbsp;</td>"
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
            INVENTORYMAC = N2Str2Zero(rsRANKFLE!MAC)
            TOTALINVENTORYONHAND = TOTALINVENTORYONHAND + N2Str2IntZero(rsRANKFLE!ONHAND)
            TOTALINVENTORYMAC = TOTALINVENTORYMAC + (INVENTORYONHAND * INVENTORYMAC)
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "1" Then
                RANKA1QTY = RANKA1QTY + INVENTORYONHAND
                RANKA1EXTCOST = RANKA1EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "2" Then
                RANKA2QTY = RANKA2QTY + INVENTORYONHAND
                RANKA2EXTCOST = RANKA2EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "A" And INVENTORYSUBCLASS = "3" Then
                RANKA3QTY = RANKA3QTY + INVENTORYONHAND
                RANKA3EXTCOST = RANKA3EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "B" Then
                RANKBQTY = RANKBQTY + INVENTORYONHAND
                RANKBEXTCOST = RANKBEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "C" Then
                RANKCQTY = RANKCQTY + INVENTORYONHAND
                RANKCEXTCOST = RANKCEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "D" Then
                RANKDQTY = RANKDQTY + INVENTORYONHAND
                RANKDEXTCOST = RANKDEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "1" Then
                RANKE1QTY = RANKE1QTY + INVENTORYONHAND
                RANKE1EXTCOST = RANKE1EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "2" Then
                RANKE2QTY = RANKE2QTY + INVENTORYONHAND
                RANKE2EXTCOST = RANKE2EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "3" Then
                RANKE3QTY = RANKE3QTY + INVENTORYONHAND
                RANKE3EXTCOST = RANKE3EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "4" Then
                RANKE4QTY = RANKE4QTY + INVENTORYONHAND
                RANKE4EXTCOST = RANKE4EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "E" And INVENTORYSUBCLASS = "5" Then
                RANKE5QTY = RANKE5QTY + INVENTORYONHAND
                RANKE5EXTCOST = RANKE5EXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "F" Then
                RANKFQTY = RANKFQTY + INVENTORYONHAND
                RANKFEXTCOST = RANKFEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            rsRANKFLE.MoveNext
        Loop
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A1 - SUPER FAST MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA1QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA1EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA1QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA1EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A2 - SUPER FAST MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA2QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA2EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA2QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA2EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK A3 - SUPER FAST MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA3QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKA3EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA3QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKA3EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK B &nbsp;&nbsp;- FAST MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKBQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKBEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKBQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKBEXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK C &nbsp;&nbsp;- MEDIUM MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKCQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKCEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKCQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKCEXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK D &nbsp;&nbsp;- SLOW MOVING PARTS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKDQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKDEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKDQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKDEXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E1 - NON MOVING PARTS FOR 1 YEAR</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE1QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE1EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE1QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE1EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E2 - NON MOVING PARTS FOR 2 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE2QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE2EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE2QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE2EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E3 - NON MOVING PARTS FOR 3 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE3QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE3EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE3QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE3EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E4 - NON MOVING PARTS FOR 4 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE4QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE4EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE4QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE4EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK E5 - NON MOVING PARTS FOR 5 YEARS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE5QTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKE5EXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE5QTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKE5EXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>RANK F &nbsp;&nbsp;- NEW ITEMS</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKFQTY, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(RANKFEXTCOST, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKFQTY / TOTALINVENTORYONHAND) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format((RANKFEXTCOST / TOTALINVENTORYMAC) * 100, "##0.00") & " %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------------<br>"
        Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
        Print #1, "<tr>"
        Print #1, "<td width=50% align=left><FONT SIZE=2>GRAND TOTAL</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(TOTALINVENTORYONHAND, "###,##0") & "</FONT></td>"
        Print #1, "<td width=15% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(TOTALINVENTORYMAC, "#,###,##0.00") & "</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>100.00 %</FONT></td>"
        Print #1, "<td width=10% align=right><FONT SIZE=2 FACE=TIMES NEW ROMAN>100.00 %</FONT></td>"
        Print #1, "</tr>"
        Print #1, "</table>"
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open PMIS_REPORT_PATH & "RANKSUM.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRank.Navigate PMIS_REPORT_PATH & "RANKSUM.HTML"
            DoEvents
            browRank.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "RANKING REPORT") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If IsDate(cboDate_Gen.Text) = True Then
        If Option1.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_RankSales where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "SALES RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "Rankings\Sales Ranking Report.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "RANKING REPORT", "", "", "", cboDate_Gen & " " & Option1.Caption, "", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option2.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "PIECE RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "Rankings\Piece Ranking Report.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "RANKING REPORT", "", "", "Parts", cboDate_Gen & " " & Option2.Caption, "", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option3.Value = True Then
            Set rsSTKSTAT = New ADODB.Recordset
            rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "SHELF STOCK RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "Rankings\Shelf Stock Ranking Report.rpt", "{STKSTAT.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "RANKING REPORT", "", "", "Parts", cboDate_Gen & " " & Option3.Caption, "", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option4.Value = True Then
            Set rsSTKSTAT = New ADODB.Recordset
            rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "COST RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "Rankings\Cost Ranking Report.rpt", "{STKSTAT.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "RANKING REPORT", "", "", "Parts", cboDate_Gen & " " & Option4.Caption, "", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option5.Value = True Then
            Set rsSTKSTAT = New ADODB.Recordset
            rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "GROSS PROFIT MARGIN RANKING REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "Rankings\Gross Profit Margin Ranking Report.rpt", "{STKSTAT.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "RANKING REPORT", "", "", "Parts", cboDate_Gen & " " & Option5.Caption, "", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option6.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                If chkSummary.Value = 1 Then
                    RANKSUMPRINTING
                    NEW_LogAudit "V", "RANKING REPORT", " ", "", "", cboDate_Gen & " " & Option6.Caption, "SUMMARY", ""
                Else
                    rptPrintRankfle.WindowTitle = "INVENTORY RANKING REPORT"
                    rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ranking.rpt", "{rankfle.type} = 'P' and {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                    NEW_LogAudit "V", "RANKING REPORT", "", "", "", cboDate_Gen & " " & Option6.Caption, "", ""
                End If
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
    Else
        MsgSpeechBox "Invalid Date Generated!"
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (RANKING REPORTS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "RANKING REPORTS", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set rsRANKFLE = New ADODB.Recordset
    rsRANKFLE.Open "select date_gen from PMIS_RankFle group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
        cboDate_Gen.Clear
        Do While Not rsRANKFLE.EOF
            cboDate_Gen.AddItem Null2Date(rsRANKFLE!DATE_GEN)
            rsRANKFLE.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PrintRankfle = Nothing
    UnloadForm Me
End Sub

