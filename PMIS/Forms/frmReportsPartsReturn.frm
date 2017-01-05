VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReports_PartsReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Return Reports"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsPartsReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   525
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
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
      Format          =   110624769
      CurrentDate     =   40066
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      Top             =   510
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
      Format          =   110624769
      CurrentDate     =   40066
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&GENERATE REPORT"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1380
      Width           =   1755
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   582
      Picture         =   "frmReportsPartsReturn.frx":1082
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmReportsPartsReturn.frx":109E
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
      Caption         =   "RETURNED PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   1755
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
      Left            =   180
      TabIndex        =   4
      Top             =   570
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
      Left            =   2430
      TabIndex        =   3
      Top             =   570
      Width           =   240
   End
End
Attribute VB_Name = "frmReports_PartsReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    If Len(Dir(App.Path & "\RET_FROM_SERVICE.XLT")) <= 0 Then
        If EXTRACT_FILES(110, "RET_FROM_SERVICE.XLT") = False Then
            MsgBox "Please Put RET_FROM_SERVICE.XLT on " & vbCrLf & App.Path, vbInformation
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
    Dim xTranDate                                      As String
    Dim COUNTER                                        As Long
    Dim XSTOCK_ORD                                     As String
    Dim XSTOCK_DESC                                    As String
    Dim XREQ_BY                                        As String
    Dim XVERI_BY                                       As String
    Dim XQTY_ISS                                       As Integer
    Dim XTYPE                                          As String
    Dim XQTY_REQ                                       As Integer
    Dim XPRICE                                         As Double
    Dim FDATETO                                        As Date
    Dim FDATEFROM                                      As Date
    Dim STATUS                                         As String
    Dim XSUMQTY_ISS                                    As Integer
    Dim XSUMQTY_REQ                                    As Integer
    Dim XSUMPRICE                                      As Double
    Dim XBALANCE                                       As Integer
    Dim XSUMBALANCE                                    As Integer
    Dim XTOTALPRICE                                     As Double
    Dim XSUMTOTALPRICE                                  As Double
    Dim rg                                             As Excel.Range
    'HEADER
    COUNTER = 10

    prgExcelGen.Text = ""

    Set xlApp = CreateObject("Excel.Application")
    'Set xlBook = xlApp.Workbooks.Open(App.Path & "\RET_FROM_SERVICE.XLT")
    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\RET_FROM_SERVICE.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(4, "C") = COMPANY_NAME
    xlSheet.Cells(5, "C") = COMPANY_ADDRESS

    FDATETO = CDate(DTPicker1)
    FDATEFROM = CDate(DTPicker2)


    Set RSHEADER = gconDMIS.Execute(GETHEADER(FDATETO, FDATEFROM))


    If Not (RSHEADER.BOF And RSHEADER.EOF) Then
          Dim lng                                            As Long
    
            lng = gconDMIS.Execute(GETCOUNT(FDATETO, FDATEFROM)).Fields(0).Value

            If lng > 0 Then
                prgExcelGen.Max = lng
                prgExcelGen.Value = 11
            End If
    
        Do While Not RSHEADER.EOF
            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"
            DoEvents
            
            XSUMQTY_ISS = 0
            XSUMQTY_REQ = 0
            XSUMPRICE = 0
            XSUMBALANCE = 0
            XSUMTOTALPRICE = 0
            
            If IsNull((RSHEADER!VERI_BY)) = True Then
                STATUS = "Not Yet Verify"
            Else
                STATUS = "Verified"
            End If
            
            XRONO = Null2String(RSHEADER!REP_OR)
            xTranDate = Null2String(RSHEADER!DATE_REQ)
            XREQ_BY = Null2String(RSHEADER!REQ_BY)
            XVERI_BY = Null2String(RSHEADER!VERI_BY)
            
       
            xlSheet.Cells(COUNTER, "A") = XRONO
            xlSheet.Cells(COUNTER, "B") = xTranDate
            xlSheet.Cells(COUNTER, "C") = XREQ_BY
            xlSheet.Cells(COUNTER, "D") = STATUS
            
            Set RSDETAILS = gconDMIS.Execute(GETDETAILS(XRONO))
            'DETAILS
            If Not (RSDETAILS.EOF And RSDETAILS.BOF) Then
                Do While Not RSDETAILS.EOF
                    
                    XQTY_ISS = 0
                    XQTY_REQ = 0
                    XPRICE = 0
                    XBALANCE = 0
                    XTOTALPRICE = 0
                    
                    XSTOCK_ORD = Trim(RSDETAILS!STOCKNO)
                    XSTOCK_DESC = Trim(RSDETAILS!STOCKDESC)
                    XTYPE = Trim(RSDETAILS!STOCK_TYPE)
                    XQTY_ISS = Trim(RSDETAILS!QTY_ISS)
                    XQTY_REQ = Trim(RSDETAILS!QTY_REQ)
                    XPRICE = Trim(RSDETAILS!TRANUPRICE)
                    
                    XSUMQTY_ISS = XSUMQTY_ISS + XQTY_ISS
                    XSUMQTY_REQ = XSUMQTY_REQ + XQTY_REQ
                    XSUMPRICE = XSUMPRICE + XPRICE
                    XBALANCE = XQTY_ISS - XQTY_REQ
                    XSUMBALANCE = XSUMBALANCE + XBALANCE
                    XTOTALPRICE = Round(XBALANCE * XPRICE, 2)
                    XSUMTOTALPRICE = XSUMTOTALPRICE + XTOTALPRICE
                    
                    xlSheet.Cells(COUNTER, "E") = XSTOCK_ORD
                    xlSheet.Cells(COUNTER, "F") = XSTOCK_DESC
                    xlSheet.Cells(COUNTER, "G") = XTYPE
                    xlSheet.Cells(COUNTER, "H") = XQTY_ISS
                    xlSheet.Cells(COUNTER, "I") = XQTY_REQ
                    xlSheet.Cells(COUNTER, "J") = XBALANCE
                    xlSheet.Cells(COUNTER, "K") = XPRICE
                    xlSheet.Cells(COUNTER, "L") = XTOTALPRICE


                    Set rg = xlSheet.Range(xlSheet.Cells(COUNTER, "A"), xlSheet.Cells(COUNTER, "L"))
                    rg.Borders.LineStyle = 1
                    COUNTER = COUNTER + 1
                    RSDETAILS.MoveNext
                Loop
            End If

            
            xlSheet.Cells(COUNTER, "H") = XSUMQTY_ISS
            xlSheet.Cells(COUNTER, "I") = XSUMQTY_REQ
            xlSheet.Cells(COUNTER, "J") = XSUMBALANCE
            xlSheet.Cells(COUNTER, "K") = XSUMPRICE
            xlSheet.Cells(COUNTER, "L") = XSUMTOTALPRICE

            Set rg = xlSheet.Range(xlSheet.Cells(COUNTER, "F"), xlSheet.Cells(COUNTER, "L"))
            rg.Font.Bold = True


            COUNTER = COUNTER + 2
            RSHEADER.MoveNext
            prgExcelGen.Value = prgExcelGen.Value + 1
        Loop
        prgExcelGen.Text = " Generation (100% Completed)"
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

        SQLTXT = "SELECT DISTINCT(REP_OR),DATE_REQ,REQ_BY,VERI_BY FROM" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.ITEMID,A.REP_OR,A.DATE_REQ,A.REQ_BY,A.VERI_BY,B.STOCKNO,B.STOCK_TYPE,B.QTY_ISS,B.QTY_REQ" & vbCrLf
        SQLTXT = SQLTXT & "FROM CSMS_RETURN_HD A INNER JOIN CSMS_RETURN_DET B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.REP_OR = B.REP_OR AND A.ID = B.ID_HD WHERE A.STATUS = 'P'" & vbCrLf
        SQLTXT = SQLTXT & ") X INNER JOIN PMIS_DAYTRAN Y ON X.ITEMID = Y.ID AND X.STOCK_TYPE = Y.[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & "WHERE DATE_REQ > = '" & XDATETO & "' AND DATE_REQ < = '" & XDATEFROM & "' ORDER BY REP_OR DESC" & vbCrLf

        GETHEADER = SQLTXT

End Function


Private Function GETCOUNT(XDATETO As Date, XDATEFROM As Date) As String
        Dim SQLTXT                                         As String

        SQLTXT = "SELECT COUNT(*) FROM(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,STOCKNO,STOCK_TYPE,QTY_ISS,QTY_REQ AS QTY_REQ," & vbCrLf
        SQLTXT = SQLTXT & "TRANUPRICE From" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.ITEMID,A.REP_OR,A.DATE_REQ,A.REQ_BY,A.VERI_BY,B.STOCKNO,B.STOCK_TYPE,B.QTY_ISS,B.QTY_REQ" & vbCrLf
        SQLTXT = SQLTXT & "FROM CSMS_RETURN_HD A INNER JOIN CSMS_RETURN_DET B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.REP_OR = B.REP_OR AND A.ID = B.ID_HD WHERE A.STATUS = 'P'" & vbCrLf
        SQLTXT = SQLTXT & ") X INNER JOIN PMIS_DAYTRAN Y" & vbCrLf
        SQLTXT = SQLTXT & "ON X.ITEMID = Y.ID AND X.STOCK_TYPE = Y.[TYPE] AND LTRIM(RTRIM(X.STOCKNO)) = LTRIM(RTRIM(Y.STOCK_ORD))" & vbCrLf
        SQLTXT = SQLTXT & ") T" & vbCrLf

        GETCOUNT = SQLTXT
End Function

Private Function GETDETAILS(REP_OR As String) As String
        Dim SQLTXT                                         As String
        
        SQLTXT = "SELECT * FROM (" & vbCrLf
        SQLTXT = SQLTXT & "SELECT T.REP_OR,T.STOCKNO,U.STOCKDESC,T.STOCK_TYPE,T.QTY_ISS,T.QTY_REQ,T.TRANUPRICE FROM(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT REP_OR,STOCKNO,STOCK_TYPE,QTY_ISS,QTY_REQ AS QTY_REQ," & vbCrLf
        SQLTXT = SQLTXT & "TRANUPRICE From" & vbCrLf
        SQLTXT = SQLTXT & "(" & vbCrLf
        SQLTXT = SQLTXT & "SELECT B.ITEMID,A.REP_OR,A.DATE_REQ,A.REQ_BY,A.VERI_BY,B.STOCKNO,B.STOCK_TYPE,B.QTY_ISS,B.QTY_REQ" & vbCrLf
        SQLTXT = SQLTXT & "FROM CSMS_RETURN_HD A INNER JOIN CSMS_RETURN_DET B" & vbCrLf
        SQLTXT = SQLTXT & "ON A.REP_OR = B.REP_OR AND A.ID = B.ID_HD WHERE A.STATUS = 'P'" & vbCrLf
        SQLTXT = SQLTXT & ") X INNER JOIN PMIS_DAYTRAN Y" & vbCrLf
        SQLTXT = SQLTXT & "ON X.ITEMID = Y.ID AND X.STOCK_TYPE = Y.[TYPE] AND LTRIM(RTRIM(X.STOCKNO)) = LTRIM(RTRIM(Y.STOCK_ORD))" & vbCrLf
        SQLTXT = SQLTXT & ") T INNER JOIN PMIS_STOCKMAS U ON T.STOCKNO = U.STOCKNO AND T.STOCK_TYPE = U.[TYPE]" & vbCrLf
        SQLTXT = SQLTXT & ") T WHERE REP_OR = '" & REP_OR & "'" & vbCrLf
    
        GETDETAILS = SQLTXT
End Function


Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    DTPicker1.Value = firstDay(LOGDATE)
    DTPicker2.Value = LOGDATE

End Sub


