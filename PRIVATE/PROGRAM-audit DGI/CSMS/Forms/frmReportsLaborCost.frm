VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReportsLaborcost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labor Cost Reports"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsLaborCost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   4755
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   525
      Left            =   3780
      TabIndex        =   5
      Top             =   1770
      Width           =   945
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Print"
      Height          =   525
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   4725
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   600
         TabIndex        =   6
         Top             =   330
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50331649
         CurrentDate     =   40249
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   2970
         TabIndex        =   7
         Top             =   330
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50331649
         CurrentDate     =   40249
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2640
         TabIndex        =   2
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   75
         TabIndex        =   1
         Top             =   390
         Width           =   480
      End
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   270
      Left            =   30
      TabIndex        =   9
      Top             =   1470
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   476
      Picture         =   "frmReportsLaborCost.frx":1082
      ForeColor       =   0
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmReportsLaborCost.frx":109E
      Max             =   200
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Labor Cost =((((Emp Monthly Salary) * 12) / 314) / 8) * hrs work"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Width           =   4245
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   30
      Picture         =   "frmReportsLaborCost.frx":10BA
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   -30
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmReportsLaborcost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    cmdView.Enabled = False
    cmdExit.Enabled = False
    
    Screen.MousePointer = 11
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim rstmp As New ADODB.Recordset
    Dim COUNTER As Long
    Dim SAME_RO    As String
    Dim RG                                             As Excel.Range
    Dim SUM_DETCOST As Double
    Dim SUM_DETAMT As Double
    Dim SUM_VAT As Double
    Dim SUM_TOTAL_AMT As Double
    Dim prgcounter As Long
    
     If Len(Dir(App.Path & "\LABORCOST.xlt")) <= 0 Then
        If EXTRACT_FILES(112, "\LABORCOST.xlt") = False Then
            MsgBox "Please Put LABORCOST.XLT on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\LABORCOST.XLT")
    Set xlSheet = xlBook.Worksheets(1)
    
  
    prgExcelGen.Text = ""
    
    prgcounter = gconDMIS.Execute(GetCount(DTPicker1, DTPicker2)).Fields(0).Value
    
    If prgcounter > 0 Then
        prgExcelGen.Max = prgcounter
        prgExcelGen.Value = 0
    End If
    
    xlSheet.Cells(1, "B") = COMPANY_NAME
    xlSheet.Cells(2, "B") = COMPANY_ADDRESS
    
    COUNTER = 6: SAME_RO = "":
   
    Set rstmp = gconDMIS.Execute(GetSQL())
    
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            DoEvents
            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"

            If SAME_RO <> Null2String(rstmp!RO_NO) Then
              
                SAME_RO = Null2String(rstmp!RO_NO)
                COUNTER = COUNTER + 2
                Set RG = xlSheet.Range(xlSheet.Cells(6, "A"), xlSheet.Cells(6, "J"))
                RG.Font.Color = vbWhite
                
                xlSheet.Cells(COUNTER - 2, "G") = SUM_DETCOST
                xlSheet.Cells(COUNTER - 2, "H") = SUM_DETAMT
                xlSheet.Cells(COUNTER - 2, "I") = SUM_VAT
                xlSheet.Cells(COUNTER - 2, "J") = SUM_TOTAL_AMT
                
                SUM_DETCOST = 0: SUM_DETAMT = 0: SUM_TOTAL_AMT = 0: SUM_VAT = 0:
                Set RG = xlSheet.Range(xlSheet.Cells(COUNTER - 2, "G"), xlSheet.Cells(COUNTER - 2, "J"))
                RG.Font.Bold = True
                                
            End If
            
            xlSheet.Cells(COUNTER, "A") = Null2String(rstmp!RO_NO)
            xlSheet.Cells(COUNTER, "B") = Null2String(rstmp!invoice)
            xlSheet.Cells(COUNTER, "C") = Null2String(rstmp!NIYM)
            xlSheet.Cells(COUNTER, "D") = Null2String(rstmp!JOBTYPE)
            xlSheet.Cells(COUNTER, "E") = Null2String(rstmp!DETCDE)
            xlSheet.Cells(COUNTER, "F") = Null2String(rstmp!DETDSC)
            xlSheet.Cells(COUNTER, "G") = NumericVal(rstmp!DetCost)
            xlSheet.Cells(COUNTER, "H") = NumericVal(rstmp!DETAMT)
            xlSheet.Cells(COUNTER, "I") = NumericVal(rstmp!VAT)
            xlSheet.Cells(COUNTER, "J") = NumericVal(rstmp!TOTAL_AMT)
            
            SUM_DETAMT = SUM_DETAMT + NumericVal(rstmp!DETAMT)
            SUM_DETCOST = SUM_DETCOST + NumericVal(rstmp!DetCost)
            SUM_VAT = SUM_VAT + NumericVal(rstmp!VAT)
            SUM_TOTAL_AMT = SUM_TOTAL_AMT + NumericVal(rstmp!TOTAL_AMT)
            
            Set RG = xlSheet.Range(xlSheet.Cells(COUNTER, "A"), xlSheet.Cells(COUNTER, "J"))
            RG.Borders.LineStyle = 1
            
            COUNTER = COUNTER + 1
            prgExcelGen.Value = prgExcelGen.Value + 1
        
        rstmp.MoveNext
        Loop
        
        If rstmp.EOF = True Then
                xlSheet.Cells(COUNTER, "G") = SUM_DETCOST
                xlSheet.Cells(COUNTER, "H") = SUM_DETAMT
                xlSheet.Cells(COUNTER, "I") = SUM_VAT
                xlSheet.Cells(COUNTER, "J") = SUM_TOTAL_AMT
                
                Set RG = xlSheet.Range(xlSheet.Cells(COUNTER, "G"), xlSheet.Cells(COUNTER, "J"))
                RG.Font.Bold = True
        End If
    Else
        MessagePop InfoFriend, "No Records", "No Records Found!"
        Screen.MousePointer = 0
        cmdView.Enabled = True
        cmdExit.Enabled = True
        Exit Sub
    End If
    
    prgExcelGen.Value = 0
    prgExcelGen.Text = "Generation (100% Completed)"
    xlApp.Visible = True
    
    cmdView.Enabled = True
    cmdExit.Enabled = True
    
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Screen.MousePointer = 0
    Set rstmp = Nothing
End Sub

Function GetCount(DATEFROM As Date, DATETO As Date) As String
    Dim SQL As String
    Dim rstmp As New ADODB.Recordset
    
    SQL = "SELECT COUNT(*) FROM CSMS_REPOR A INNER JOIN CSMS_RO_DET B ON A.REP_OR = B.REP_OR" & vbCrLf
    SQL = SQL & "AND A.TRANSTYPE = B.TRANSTYPE WHERE  B.LIVIL = '1' AND A.INVOICE IS NOT NULL AND A.TRANSTYPE = 'R'" & vbCrLf
    SQL = SQL & "AND A.DTE_REL > ='" & DATEFROM & "' AND A.DTE_REL < = '" & DATETO & "'" & vbCrLf
    
    GetCount = SQL
End Function

Function GetSQL() As String
    Dim SQL As String


    SQL = "SELECT RO_NO,INVOICE,NIYM,JOBTYPE,DETCDE,DETDSC," & vbCrLf
    SQL = SQL & "CAST(CASE ROTYPE WHEN 'SR' THEN S_DETCOST" & vbCrLf
    SQL = SQL & "ELSE (CASE WHEN ISNULL(SALARY,0) = 0 THEN" & vbCrLf
    SQL = SQL & "(CASE WHEN S_DETCOST < 0 THEN 0 ELSE S_DETCOST END )" & vbCrLf
    SQL = SQL & "ELSE (SALARY * (CASE WHEN J_HRSWRK > B_HRSWRK THEN  J_HRSWRK ELSE B_HRSWRK END)) END)" & vbCrLf
    SQL = SQL & "END AS DECIMAL(18,2)) AS DETCOST,DETAMT , VAT, TOTAL_AMT" & vbCrLf
    SQL = SQL & "From" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT RO_NO,INVOICE,NIYM,JOBTYPE,DETCDE,DETDSC,ISNULL(ROTYPE,'') AS ROTYPE,SUM(J_HRSWRK) AS J_HRSWRK,B_HRSWRK," & vbCrLf
    SQL = SQL & "ISNULL(SALARY,0) AS SALARY,CAST(ISNULL(DETCOST,0) AS DECIMAL(18,2)) as S_DETCOST,CAST(DETAMT AS DECIMAL(18,2)) AS DETAMT , VAT, TOTAL_AMT,DTE_REL, DTE_COMP" & vbCrLf
    SQL = SQL & "From" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT X.TECHNICIAN AS EMPNO,X.TECH_NAME,ISNULL(Y.HRSWRK,0) AS B_HRSWRK,ISNULL(X.HrsWorked,0) AS J_HRSWRK,X.TECHCODE,X.RO_NO,X.DETCDE,Y.DETDSC,Y.JOBTYPE,Y.DISCOUNT_2,Y.ROTYPE,Y.DTE_REL,Y.DTE_COMP," & vbCrLf
    SQL = SQL & "(SELECT CAST((((SALARY*12)/314)/8) AS DECIMAL(18,2)) FROM" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT A.EMPNO," & vbCrLf
    SQL = SQL & "CASE ISNULL(A.SALARYCODE,'0') WHEN '0' THEN B.SALARY" & vbCrLf
    SQL = SQL & "Else" & vbCrLf
    SQL = SQL & "(CASE  ISNULL(B.DAILYRATE,0) WHEN 0 THEN B.SALARY ELSE B.SALARY END) END AS SALARY," & vbCrLf
    SQL = SQL & "SUBSTRING(LASTNAME,1,1) + SUBSTRING(FIRSTNAME,1,1) + SUBSTRING(MIDDLENAME,1,1) AS TECHCODE" & vbCrLf
    SQL = SQL & "FROM HRMS_EMPINFO A LEFT OUTER JOIN HRMS_SALARYGRADE B ON A.SALARYCODE = CODE WHERE IS_TECHNICIAN =1" & vbCrLf
    SQL = SQL & ") T WHERE T.EMPNO = X.TECHNICIAN) AS SALARY,Y.VAT,Y.TOTAL_AMT,Y.DETAMT,Y.DETCOST,Y.INVOICE,Y.NIYM FROM CSMS_JOBCLOCK X" & vbCrLf
    SQL = SQL & "LEFT OUTER JOIN" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT A.REP_OR,INVOICE,B.JOBTYPE,B.DETCDE,DETDSC,ISNULL(DETAMT,0) AS DETAMT ,A.DTE_REL,A.DTE_COMP," & vbCrLf
    SQL = SQL & "CAST((ISNULL(DETAMT,0) * 0.12) AS DECIMAL(18,2)) AS VAT,A.NIYM,B.LIVIL," & vbCrLf
    SQL = SQL & "CAST(ISNULL(DETAMT,0) + (ISNULL(DETAMT,0) * 0.12) AS DECIMAL(18,2)) AS TOTAL_AMT," & vbCrLf
    SQL = SQL & "A.TRANSTYPE,ISNULL(B.HRSWRK,0) AS HRSWRK,ISNULL(B.DISCOUNT_2,0) AS DISCOUNT_2 ,B.ROTYPE,B.DETCOST" & vbCrLf
    SQL = SQL & "From" & vbCrLf
    SQL = SQL & "CSMS_REPOR A INNER JOIN CSMS_RO_DET B ON A.REP_OR = B.REP_OR AND A.TRANSTYPE = B.TRANSTYPE" & vbCrLf
    SQL = SQL & ") Y ON X.RO_NO = Y.REP_OR AND X.DETCDE = Y.DETCDE WHERE Y.TRANSTYPE = 'R' AND Y.INVOICE IS NOT NULL AND Y.LIVIL = '1'" & vbCrLf
    SQL = SQL & ") RPT_LABORCOST GROUP BY RO_NO,INVOICE,NIYM,JOBTYPE,DETCDE,DETDSC,ROTYPE,B_HRSWRK,DETAMT ,VAT, TOTAL_AMT,SALARY,DETCOST,DTE_REL, DTE_COMP" & vbCrLf
    SQL = SQL & ") T WHERE DTE_COMP > = '" & DTPicker1 & "' AND DTE_COMP < = '" & DTPicker2 & "' ORDER BY RO_NO DESC" & vbCrLf
            
    GetSQL = SQL
End Function

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    DTPicker1.Value = firstDay(LOGDATE)
    DTPicker2.Value = LOGDATE
End Sub

