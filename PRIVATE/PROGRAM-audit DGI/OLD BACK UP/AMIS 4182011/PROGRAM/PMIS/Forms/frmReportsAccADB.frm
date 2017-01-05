VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportsAccADB 
   Caption         =   "ACCESSORIES"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   3000
      TabIndex        =   6
      Top             =   1020
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   609
      _Version        =   393216
      Format          =   49938433
      CurrentDate     =   40066
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   990
      TabIndex        =   5
      Top             =   1020
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      Format          =   49938433
      CurrentDate     =   40066
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&GENERATE REPORT"
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   1980
      Width           =   1755
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   390
      TabIndex        =   1
      Top             =   1500
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   582
      Picture         =   "frmReportsAccADB.frx":0000
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmReportsAccADB.frx":001C
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "TO"
      Height          =   405
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label 
      Caption         =   "FROM"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1050
      Width           =   915
   End
   Begin VB.Label Label 
      Caption         =   "ADVANCE BILL VS ADVANCE BILL ISSUANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   5085
   End
End
Attribute VB_Name = "frmReportsAccADB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    If Len(Dir(App.Path & "\ADB.XLT")) <= 0 Then
        If EXTRACT_FILES(108, "ADB.XLT") = False Then
            MsgBox "Please Put ADB.XLT on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim SQL                                            As String
    Dim RSHEADER                                       As New ADODB.Recordset
    Dim RSDETAILS                                      As New ADODB.Recordset
    Dim XRONO                                          As String
    Dim XTRANNO                                        As String
    Dim XTRANDATE                                      As String
    Dim COUNTER                                        As Long
    Dim XSTOCK_ORD                                     As String
    Dim XSALES_ORIGIN                                  As String
    Dim XFILL                                          As Integer
    Dim XSUMPRICE                                      As Double
    Dim XONHAND                                        As Integer
    Dim XTRANQTY                                       As Integer
    Dim XBALANCE                                       As Integer
    Dim RG                                             As Excel.Range
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

    prgExcelGen.Max = 72
    prgExcelGen.Value = 0
    Set RSHEADER = gconDMIS.Execute(GETHEADER(FDATETO, FDATEFROM))
    If Not (RSHEADER.BOF And RSHEADER.EOF) Then
        RSHEADER.MoveFirst
        Do While Not RSHEADER.EOF

            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"
            DoEvents
            XSUMTRANQTY = 0
            XSUMTRANUPRICE = 0
            XSUMFILL = 0
            XSUMBALANCE = 0
            XSUMONHAND = 0

            XRONO = Trim(RSHEADER!RoNo)
            XTRANNO = Trim(RSHEADER!TRANNO)
            XTRANDATE = Trim(RSHEADER!TRANDATE)
            XSALES_ORIGIN = Trim(RSHEADER!SALES_ORIGIN)

            xlSheet.Cells(COUNTER, "A") = XTRANNO
            xlSheet.Cells(COUNTER, "B") = XRONO
            xlSheet.Cells(COUNTER, "C") = XTRANDATE
            xlSheet.Cells(COUNTER, "D") = XSALES_ORIGIN

            Set RSDETAILS = gconDMIS.Execute(GETDETAILS(XRONO))
            'DETAILS
            If Not (RSDETAILS.EOF And RSDETAILS.BOF) Then
                RSDETAILS.MoveFirst
                Do While Not RSDETAILS.EOF

                    XSTOCK_ORD = Trim(RSDETAILS!STOCK_ORD)
                    XONHAND = Trim(RSDETAILS!ONHAND)
                    XTRANQTY = Trim(RSDETAILS!tranqty)
                    XTRANUPRICE = Trim(RSDETAILS!TRANUPRICE)
                    XFILL = GetTotal_ADB_Filled(XRONO, XSTOCK_ORD)
                    XBALANCE = XTRANQTY - XFILL

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

                    COUNTER = COUNTER + 1
                    RSDETAILS.MoveNext
                Loop
            End If

            xlSheet.Cells(COUNTER, "F") = XSUMONHAND
            xlSheet.Cells(COUNTER, "G") = XSUMTRANQTY
            xlSheet.Cells(COUNTER, "H") = XSUMFILL
            xlSheet.Cells(COUNTER, "I") = XSUMBALANCE
            xlSheet.Cells(COUNTER, "J") = XSUMTRANUPRICE

            Set RG = xlSheet.Range(xlSheet.Cells(COUNTER, "F"), xlSheet.Cells(COUNTER, "J"))
            RG.Font.Bold = True

            COUNTER = COUNTER + 2
            RSHEADER.MoveNext
            prgExcelGen.Value = prgExcelGen.Value + 1
        Loop
    Else
        ShowNoRecord
    End If
    prgExcelGen.Text = " GENERATING REPORTS (100% Completed)"
    xlApp.Visible = True
    Set xlApp = Nothing

End Sub


Private Function GETHEADER(XDATETO As Date, XDATEFROM As Date) As String
    Dim sqltxt                                         As String

    sqltxt = "SELECT "
    sqltxt = sqltxt & "Distinct (T.RoNo) as RONO, (T.TRANNO) as TRANNO, T.TRANDATE as TRANDATE,T.SALES_ORIGIN AS SALES_ORIGIN "
    sqltxt = sqltxt & "From( "
    sqltxt = sqltxt & "SELECT TRANDATE,TRANNO ,RONO ,'CURRT' AS DSTATUS ,SALES_ORIGIN "
    sqltxt = sqltxt & "FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='A' AND ISNULL(STATUS3,'')  <>'F' AND "
    sqltxt = sqltxt & "ISNULL(STATUS2,'') <>'R' AND (STATUS='P' OR STATUS='B') "
    sqltxt = sqltxt & "Union "
    sqltxt = sqltxt & "SELECT "
    sqltxt = sqltxt & "TRANDATE,TRANNO ,RONO ,'HIST' AS DSTATUS ,SALES_ORIGIN "
    sqltxt = sqltxt & "FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='A' AND "
    sqltxt = sqltxt & "ISNULL(STATUS3,'') <>'F' AND  ISNULL(STATUS2,'')  <>'R' AND "
    sqltxt = sqltxt & "(STATUS='P' OR STATUS='B')) T INNER JOIN PMIS_ALLDAYTRAN Y "
    sqltxt = sqltxt & "ON T.TRANNO = Y.TRANNO WHERE [TYPE] = 'P' AND TRANTYPE = 'ADB' AND SALES_ORIGIN = 'S' "
    sqltxt = sqltxt & "AND T.TRANDATE > = '" & XDATETO & "' AND T.TRANDATE < = '" & XDATEFROM & "'"
    sqltxt = sqltxt & "ORDER BY TRANDATE ASC"

    GETHEADER = sqltxt
End Function

Private Function GETDETAILS(REP_OR As String) As String
    Dim sqltxt                                         As String


    sqltxt = sqltxt & "SELECT STOCK_ORD ,AVG(PMIS_STOCKMAS.ONHAND) ONHAND , sum(TRANQTY) as TRANQTY,TRANUPRICE FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS "
    sqltxt = sqltxt & " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO "
    sqltxt = sqltxt & " WHERE  PMIS_STOCKMAS.TYPE='A' AND "
    sqltxt = sqltxt & "(TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='A' AND RONO='" & REP_OR & "'   AND (STATUS='P' OR STATUS='B')) "
    sqltxt = sqltxt & "OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='A' AND RONO='" & REP_OR & "'   AND (STATUS='P' OR STATUS='B'))) "
    sqltxt = sqltxt & "AND TRANTYPE='ADB' GROUP BY STOCK_ORD,TRANUPRICE  ORDER BY STOCK_ORD,TRANUPRICE"

    GETDETAILS = sqltxt
End Function

Function GetTotal_ADB_Filled(xro_no As String, x_stockno As String) As Long
    Dim STR_SQLX                                       As String

    STR_SQLX = " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_TDAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HD ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TYPE=PMIS_TDAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANTYPE=PMIS_TDAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HD.TRANTYPE='RIV' AND  PMIS_ORD_HD.RONO='" & xro_no & "' AND PMIS_TDAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='P' OR PMIS_ORD_HD.STATUS='B')  AND PMIS_ORD_HD.TYPE='A' AND PMIS_ORD_HD.STATUS2='R' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='RIV' AND  PMIS_ORD_HIST.RONO='" & xro_no & "' AND PMIS_DAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='P' OR PMIS_ORD_HIST.STATUS='B')  AND PMIS_ORD_HIST.TYPE='A' AND PMIS_ORD_HIST.STATUS2='R'"
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"

    Dim RSTOTAL_FILLED                                 As ADODB.Recordset
    Set RSTOTAL_FILLED = gconDMIS.Execute(STR_SQLX)
    If Not RSTOTAL_FILLED.EOF Or Not RSTOTAL_FILLED.BOF Then
        GetTotal_ADB_Filled = N2Str2Zero(RSTOTAL_FILLED!tranqty)
    End If

End Function

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub



