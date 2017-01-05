VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReports_AppliedADB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Bill "
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   Icon            =   "frmReportsPartsADB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "APPLIED ADVACE BILL REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   150
      TabIndex        =   6
      Top             =   60
      Width           =   4515
      Begin VB.OptionButton OPT_MATERIALS 
         Caption         =   "MATERIALS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton OPT_PARTS 
         Caption         =   "PARTS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   3150
      TabIndex        =   2
      Top             =   855
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   51904513
      CurrentDate     =   40066
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   750
      TabIndex        =   1
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   51904513
      CurrentDate     =   40066
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&GENERATE REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2910
      TabIndex        =   0
      Top             =   1650
      Width           =   1755
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   150
      TabIndex        =   5
      Top             =   1260
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   582
      Picture         =   "frmReportsPartsADB.frx":6852
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmReportsPartsADB.frx":686E
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2790
      TabIndex        =   3
      Top             =   900
      Width           =   240
   End
End
Attribute VB_Name = "frmReports_AppliedADB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOCAL_STOCKTYPE                                    As String
Sub SETSTOCKSTYPE(xxx As String)
    LOCAL_STOCKTYPE = xxx
End Sub
Private Sub cmdPrint_Click()
    If Len(Dir(App.Path & "\ADB.XLT")) <= 0 Then
        If EXTRACT_FILES(108, "ADB.XLT") = False Then
            MsgBox "Please Put ADB.XLT on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim SQL                                            As String
    Dim RSHEADER                                       As New ADODB.Recordset
    Dim RSDETAILS                                      As New ADODB.Recordset
    Dim XRONO                                          As String
    Dim XTRANNO                                        As String
    Dim xTranDate                                      As String
    Dim COUNTER                                        As Long
    Dim XSTOCK_ORD                                     As String
    Dim XSALES_ORIGIN                                  As String
    Dim XFILL                                          As Integer
    Dim XSUMPRICE                                      As Double
    Dim XONHAND                                        As Integer
    Dim XTRANQTY                                       As Integer
    Dim XBALANCE                                       As Integer
    Dim rg                                             As Excel.Range
    Dim XSUMTRANQTY                                    As Integer
    Dim XTRANUPRICE                                    As Double
    Dim XSUMFILL                                       As Integer
    Dim XSUMBALANCE                                    As Integer
    Dim XSUMONHAND                                     As Integer
    Dim XSUMTRANUPRICE                                 As Double
    Dim FDATETO                                        As Date
    Dim FDATEFROM                                      As Date
    'HEADER
    COUNTER = 8

    prgExcelGen.Text = ""

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\ADB.XLT")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(3, "B") = COMPANY_NAME
    xlSheet.Cells(4, "B") = COMPANY_ADDRESS

    FDATETO = CDate(DTPicker1)
    FDATEFROM = CDate(DTPicker2)


    Dim lng                                            As Long
    lng = gconDMIS.Execute(GETCOUNT(FDATETO, FDATEFROM)).Fields(0).Value

    If lng > 0 Then
        prgExcelGen.Max = lng
        prgExcelGen.Value = 0
    End If

    Set RSHEADER = gconDMIS.Execute(GETHEADER(FDATETO, FDATEFROM))
 

    If Not (RSHEADER.BOF And RSHEADER.EOF) Then
        Do While Not RSHEADER.EOF
            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"
            DoEvents
            XSUMTRANQTY = 0
            XSUMTRANUPRICE = 0
            XSUMFILL = 0
            XSUMBALANCE = 0
            XSUMONHAND = 0

            XRONO = Trim(RSHEADER!RONO)
            XTRANNO = Trim(RSHEADER!TRANNO)
            xTranDate = Trim(RSHEADER!TRANDATE)
            XSALES_ORIGIN = Trim(RSHEADER!SALES_ORIGIN)

            xlSheet.Cells(COUNTER, "A") = XTRANNO
            xlSheet.Cells(COUNTER, "B") = XRONO
            xlSheet.Cells(COUNTER, "C") = xTranDate
            xlSheet.Cells(COUNTER, "D") = XSALES_ORIGIN
            
            Set RSDETAILS = gconDMIS.Execute(GETDETAILS(XRONO))
            'DETAILS
            If Not (RSDETAILS.EOF And RSDETAILS.BOF) Then
                Do While Not RSDETAILS.EOF
                    XSTOCK_ORD = Trim(RSDETAILS!STOCK_ORD) & " "
                    XONHAND = Trim(N2Str2IntZero(RSDETAILS!ONHAND))
                    XTRANQTY = Trim(RSDETAILS!TRANQTY)
                    XTRANUPRICE = Trim(RSDETAILS!TRANUPRICE)
                    XFILL = GetTotal_ADB_Filled(XRONO, XSTOCK_ORD)
                    XBALANCE = XTRANQTY - XFILL
                    
                    If XFILL <> 0 Then
                    Set rg = xlSheet.Range(xlSheet.Cells(COUNTER, "E"), xlSheet.Cells(COUNTER, "J"))
                    rg.Interior.ColorIndex = 6
                    End If
                    XSUMTRANQTY = XSUMTRANQTY + XTRANQTY
                    XSUMFILL = XSUMFILL + XFILL
                    XSUMTRANUPRICE = XSUMTRANUPRICE + XTRANUPRICE
                    XSUMBALANCE = XSUMBALANCE + XBALANCE
                    XSUMONHAND = XSUMONHAND + XONHAND

                    xlSheet.Cells(COUNTER, "E") = XSTOCK_ORD
                    xlSheet.Cells(COUNTER, "F") = XONHAND
                    xlSheet.Cells(COUNTER, "G") = XTRANQTY
                    xlSheet.Cells(COUNTER, "H") = XFILL
                    xlSheet.Cells(COUNTER, "I") = XBALANCE
                    xlSheet.Cells(COUNTER, "J") = XTRANUPRICE


                    Set rg = xlSheet.Range(xlSheet.Cells(COUNTER, "A"), xlSheet.Cells(COUNTER, "J"))
                    rg.Borders.LineStyle = 1
                    prgExcelGen.Value = prgExcelGen.Value + 1
                    COUNTER = COUNTER + 1
                    RSDETAILS.MoveNext
                Loop
            End If

            xlSheet.Cells(COUNTER, "F") = XSUMONHAND
            xlSheet.Cells(COUNTER, "G") = XSUMTRANQTY
            xlSheet.Cells(COUNTER, "H") = XSUMFILL
            xlSheet.Cells(COUNTER, "I") = XSUMBALANCE
            xlSheet.Cells(COUNTER, "J") = XSUMTRANUPRICE

            Set rg = xlSheet.Range(xlSheet.Cells(COUNTER, "F"), xlSheet.Cells(COUNTER, "J"))
            rg.Font.Bold = True



            COUNTER = COUNTER + 2
            RSHEADER.MoveNext
            
        Loop
        prgExcelGen.Text = "Generation (100% Completed)"
        xlApp.Visible = True

    Else
        ShowNoRecord
    End If
    Set xlApp = Nothing
    cmdPrint.Enabled = True
    Screen.MousePointer = 0
End Sub


Private Function GETHEADER(XDATETO As Date, XDATEFROM As Date) As String
    Dim SQLTXT                                         As String

    SQLTXT = "SELECT " & vbCrLf
    SQLTXT = SQLTXT & "Distinct (T.RoNo) as RONO, (T.TRANNO) as TRANNO, T.TRANDATE as TRANDATE,T.SALES_ORIGIN AS SALES_ORIGIN " & vbCrLf
    SQLTXT = SQLTXT & "From( " & vbCrLf
    SQLTXT = SQLTXT & "SELECT TRANDATE,TRANNO ,RONO ,'CURRT' AS DSTATUS ,SALES_ORIGIN " & vbCrLf
    SQLTXT = SQLTXT & "FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND (STATUS='P' OR STATUS='B') " & vbCrLf
    'sqltxt = sqltxt & " AND ISNULL(STATUS3,'')  <>'F' AND  ISNULL(STATUS2,'') <>'R' " & vbCrLf
    SQLTXT = SQLTXT & "Union " & vbCrLf
    SQLTXT = SQLTXT & "SELECT " & vbCrLf
    SQLTXT = SQLTXT & "TRANDATE,TRANNO ,RONO ,'HIST' AS DSTATUS ,SALES_ORIGIN " & vbCrLf
    SQLTXT = SQLTXT & "FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND " & vbCrLf
    'sqltxt = sqltxt & "ISNULL(STATUS3,'') <>'F' AND  ISNULL(STATUS2,'')  <>'R' AND " & vbCrLf
    SQLTXT = SQLTXT & "(STATUS='P' OR STATUS='B')) T INNER JOIN PMIS_ALLDAYTRAN Y " & vbCrLf
    SQLTXT = SQLTXT & "ON T.TRANNO = Y.TRANNO WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND TRANTYPE = 'ADB' AND SALES_ORIGIN = 'S' " & vbCrLf
    SQLTXT = SQLTXT & "AND T.TRANDATE > = '" & XDATETO & "' AND T.TRANDATE < = '" & XDATEFROM & "'" & vbCrLf
    SQLTXT = SQLTXT & "ORDER BY TRANDATE ASC"

    GETHEADER = SQLTXT

End Function


Private Function GETCOUNT(XDATETO As Date, XDATEFROM As Date) As String
    Dim SQLTXT                                         As String

    SQLTXT = "SELECT "
    SQLTXT = SQLTXT & " COUNT(*) "
    SQLTXT = SQLTXT & "From( "
    SQLTXT = SQLTXT & "SELECT TRANDATE,TRANNO ,RONO ,'CURRT' AS DSTATUS ,SALES_ORIGIN "
    SQLTXT = SQLTXT & "FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND ISNULL(STATUS3,'')  <>'F' AND "
    SQLTXT = SQLTXT & "ISNULL(STATUS2,'') <>'R' AND (STATUS='P' OR STATUS='B') "
    SQLTXT = SQLTXT & "Union "
    SQLTXT = SQLTXT & "SELECT "
    SQLTXT = SQLTXT & "TRANDATE,TRANNO ,RONO ,'HIST' AS DSTATUS ,SALES_ORIGIN "
    SQLTXT = SQLTXT & "FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND "
    SQLTXT = SQLTXT & "ISNULL(STATUS3,'') <>'F' AND  ISNULL(STATUS2,'')  <>'R' AND "
    SQLTXT = SQLTXT & "(STATUS='P' OR STATUS='B')) T INNER JOIN PMIS_ALLDAYTRAN Y "
    SQLTXT = SQLTXT & "ON T.TRANNO = Y.TRANNO WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND TRANTYPE = 'ADB' AND SALES_ORIGIN = 'S' "
    SQLTXT = SQLTXT & "AND T.TRANDATE > = '" & XDATETO & "' AND T.TRANDATE < = '" & XDATEFROM & "'"


    GETCOUNT = SQLTXT
End Function

Private Function GETDETAILS(REP_OR As String) As String
    Dim SQLTXT                                         As String


    SQLTXT = SQLTXT & "SELECT STOCK_ORD ,AVG(PMIS_STOCKMAS.ONHAND) ONHAND , sum(TRANQTY) as TRANQTY,TRANUPRICE FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS "
    SQLTXT = SQLTXT & " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO "
    SQLTXT = SQLTXT & " WHERE  PMIS_STOCKMAS.TYPE='" & LOCAL_STOCKTYPE & "' AND "
    SQLTXT = SQLTXT & "(TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & REP_OR & "'   AND (STATUS='P' OR STATUS='B')) "
    SQLTXT = SQLTXT & "OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & REP_OR & "'   AND (STATUS='P' OR STATUS='B'))) "
    SQLTXT = SQLTXT & "AND TRANTYPE='ADB' GROUP BY STOCK_ORD,TRANUPRICE  ORDER BY STOCK_ORD,TRANUPRICE"

    GETDETAILS = SQLTXT
End Function

Function GetTotal_ADB_Filled(xro_no As String, x_stockno As String) As Long
    Dim STR_SQLX                                       As String

    STR_SQLX = " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_TDAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HD ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TYPE=PMIS_TDAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANTYPE=PMIS_TDAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HD.TRANTYPE='RIV' AND  PMIS_ORD_HD.RONO='" & xro_no & "' AND PMIS_TDAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='P' OR PMIS_ORD_HD.STATUS='B')  AND PMIS_ORD_HD.TYPE='" & LOCAL_STOCKTYPE & "' AND PMIS_ORD_HD.STATUS2='R' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='RIV' AND  PMIS_ORD_HIST.RONO='" & xro_no & "' AND PMIS_DAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='P' OR PMIS_ORD_HIST.STATUS='B')  AND PMIS_ORD_HIST.TYPE='" & LOCAL_STOCKTYPE & "' AND PMIS_ORD_HIST.STATUS2='R'"
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"

    Dim RSTOTAL_FILLED                                 As ADODB.Recordset
    Set RSTOTAL_FILLED = gconDMIS.Execute(STR_SQLX)
    If Not RSTOTAL_FILLED.EOF Or Not RSTOTAL_FILLED.BOF Then
        GetTotal_ADB_Filled = N2Str2Zero(RSTOTAL_FILLED!TRANQTY)
    Else
        GetTotal_ADB_Filled = 0
    End If

End Function

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    DTPicker1.Value = firstDay(LOGDATE)
    DTPicker2.Value = LOGDATE

    If LOCAL_STOCKTYPE = "P" Then
        Me.Caption = "PARTS APPLIED ADVANCE BILL REPORT"
    ElseIf LOCAL_STOCKTYPE = "M" Then
        Me.Caption = "MATERIALS APPLIED ADVANCE BILL REPORT"
    End If

End Sub


Private Sub OPT_PARTS_Click()
    SETSTOCKSTYPE ("P")
End Sub

Private Sub OPT_MATERIALS_Click()
SETSTOCKSTYPE ("M")
End Sub
