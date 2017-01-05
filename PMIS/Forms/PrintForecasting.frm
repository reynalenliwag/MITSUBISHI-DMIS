VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_PrintForeCasting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print ForeCasting Report"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "PrintForecasting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   4425
   Begin VB.OptionButton Option8 
      Caption         =   "Seasonality Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   14
      Top             =   2820
      Width           =   3795
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Suggested Order Quantity Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   6
      Top             =   2436
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -330
      TabIndex        =   13
      Top             =   3180
      Width           =   5355
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Safety Stock Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   5
      Top             =   2055
      Width           =   3795
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Mean Absolute Deviation Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   4
      Top             =   1674
      Width           =   3795
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Linear Regression Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   3
      Top             =   1293
      Width           =   3795
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Moving Median Reports (6 Mos.)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   2
      Top             =   912
      Width           =   3795
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Six Months Moving Average Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   1
      Top             =   531
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Level Of Service Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
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
      TabIndex        =   7
      ToolTipText     =   "Choose date from list"
      Top             =   3420
      Width           =   2205
   End
   Begin SHDocVwCtl.WebBrowser browRank 
      Height          =   3945
      Left            =   6180
      TabIndex        =   11
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
      Location        =   ""
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
      TabIndex        =   8
      Top             =   3870
      Width           =   2205
   End
   Begin Crystal.CrystalReport rptPrintRankfle 
      Left            =   3180
      Top             =   4470
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
      Left            =   2040
      MouseIcon       =   "PrintForecasting.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "PrintForecasting.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   4320
      Width           =   735
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
      MouseIcon       =   "PrintForecasting.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "PrintForecasting.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   4320
      Width           =   735
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
      TabIndex        =   12
      Top             =   3450
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_PrintForeCasting"
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
    rsRANKFLE.Open "select month_gen,onhand,mac,invclass,subinvclas,date_gen from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "' order by invclass,subinvclas asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            If INVENTORYCLASS = "B" And INVENTORYSUBCLASS = "" Then
                RANKBQTY = RANKBQTY + INVENTORYONHAND
                RANKBEXTCOST = RANKBEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "C" And INVENTORYSUBCLASS = "" Then
                RANKCQTY = RANKCQTY + INVENTORYONHAND
                RANKCEXTCOST = RANKCEXTCOST + (INVENTORYMAC * INVENTORYONHAND)
            End If
            If INVENTORYCLASS = "D" And INVENTORYSUBCLASS = "" Then
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
            If INVENTORYCLASS = "F" And INVENTORYSUBCLASS = "" Then
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
    On Error GoTo ErrorCode:
    If Option8.Value = True Then
        frmPMISReports_Seasonality.Show
        frmPMISReports_Seasonality.ZOrder 0
        Exit Sub
    End If

    If IsDate(cboDate_Gen.Text) = True Then
        If Option1.Value = True Then
            Set rsSTKSTAT = New ADODB.Recordset
            rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "LEVEL OF SERVICE REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Level of Service Report.rpt", "{STKSTAT.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Level of Service Report", ""
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
                rptPrintRankfle.WindowTitle = "PAST SIX MONTHS MOVING AVERAGE REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Six Months Moving Average Report.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Parts Six Months Moving Average Report", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option3.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "MOVING MEDIAN REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Moving Median Report.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Moving Median Report", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option4.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_RankFle where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "LINEAR REGRESSION REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Linear Regression Report.rpt", "{rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Linear Regression Report", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option5.Value = True Then
            Set rsSTKSTAT = New ADODB.Recordset
            rsSTKSTAT.Open "select * from PMIS_RANKFLE where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "MEAN ABSOLUTE DEVIATION REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Mean Absolute Deviation Report.rpt", "{RANKFLE.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Mean Absolute Deviation Report", ""
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
                rptPrintRankfle.WindowTitle = "SAFETY STOCK REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\SafetyStock.rpt", "{StkStat.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Safety Stock Report", ""
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "Not Yet Generated!"
            End If
        End If
        If Option7.Value = True Then
            Set rsRANKFLE = New ADODB.Recordset
            rsRANKFLE.Open "select * from PMIS_Demand_Forecast where Date_Gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
                Screen.MousePointer = 11
                rptPrintRankfle.WindowTitle = "SUGGESTED ORDER QUANTITY REPORT"
                rptPrintRankfle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPrintRankfle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPrintRankfle, PMIS_REPORT_PATH & "ForeCasting\Suggested_Order_Qty.rpt", "{PMIS_Demand_Forecast.Date_Gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
                NEW_LogAudit "V", "FORCASTING REPORT", "", "", "Parts", cboDate_Gen, "Suggested Order Quantity Report", ""
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (FORCASTING REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "FORCASTING REPORT", "PRINTING")
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
    If FORECASTING_BUTTON_CLICK = 1 Then Option1.Value = True    'Level Of Service Report
    If FORECASTING_BUTTON_CLICK = 2 Then Option2.Value = True    'Six Months Moving Average Report
    If FORECASTING_BUTTON_CLICK = 3 Then Option3.Value = True    'Moving Median Reports(6 Mos.)
    If FORECASTING_BUTTON_CLICK = 4 Then Option4.Value = True    'Linear Regression Report
    If FORECASTING_BUTTON_CLICK = 5 Then Option5.Value = True    'Mean Absolute Deviation Report
    If FORECASTING_BUTTON_CLICK = 6 Then Option6.Value = True    'Safety Stock Report
    If FORECASTING_BUTTON_CLICK = 7 Then Option7.Value = True    'Suggested Order Quantity Report
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PrintRankfle = Nothing
    UnloadForm Me
End Sub

