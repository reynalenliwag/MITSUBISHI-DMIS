VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmBIRExtract 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Data Extraction"
   ClientHeight    =   1530
   ClientLeft      =   135
   ClientTop       =   750
   ClientWidth     =   6090
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "BIRExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6090
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   720
      Width           =   1965
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   1110
      Width           =   1965
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   5010
      Picture         =   "BIRExtract.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   690
      Width           =   945
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   765
      Left            =   4080
      Picture         =   "BIRExtract.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   945
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   330
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "BIRExtract.frx":0EDE
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "BIRExtract.frx":0EFA
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "frmBIRExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gconBIRData As ADODB.Connection

Sub Check_Click()
Dim vStartDate As String
Dim vEndDate As String

'vStartDate = N2Date2Null(txtFrom.Text)
'vEndDate = N2Date2Null(txtTo.Text)

Const TIN = "'005532084'"
Dim vReference As String
Dim vRectype As String
Dim vTrantype As String
Dim vCust_Tin As String
Dim vRef_Name As String
Dim vCust_Name As String
Dim vTrandate As String
Dim vGrs_Exempt As String
Dim vGrs_Zero_R As String
Dim vCust_Addr1 As String
Dim vCust_Addr2 As String
Dim vCustCode As String
Dim vTIN_Owner As String

vReference = "'HEADER'"
vRectype = "'H'"
vTrantype = "'S'"
vCust_Tin = TIN
vRef_Name = "'CALEB MOTOR CORPORATION'"
vCust_Name = "'CALEB MOTOR CORPORATION'"
vTrandate = vEndDate
vGrs_Exempt = 0
vGrs_Zero_R = 0
vCust_Addr1 = "'ROXAS AVENUE, DIVERSION ROAD'"
vCust_Addr2 = "'CONCEPCION PEQUEÑA, NAGA CITY'"

Dim vAddress As String
Dim vLastName As String
Dim vFirstName As String
Dim vMidName As String


Dim i As Integer
gconBIRData.Execute "Delete From BIR.DBF"

gconBIRData.Execute "Insert Into BIR.DBF " & _
                    "(Reference,Rectype,Trantype,Cust_Tin,Cust_Name,Trandate,Grs_Exempt,Grs_Zero_R,Cust_Addr1,Cust_Addr2)" & _
                    " Values (" & vReference & "," & vRectype & "," & vTrantype & "," & vCust_Tin & "," & vCust_Name & "," & vTrandate & "," & vGrs_Exempt & "," & vGrs_Zero_R & "," & vCust_Addr1 & "," & vCust_Addr2 & ")"

Dim rsORD_HIST As ADODB.Recordset
Dim rsCustomer As ADODB.Recordset

Set rsORD_HIST = New ADODB.Recordset
Set rsORD_HIST = gconPMIOS.Execute("Select * From ORD_Hist Where (trantype = 'CSH' or trantype = 'CHG') and status = 'P' and TranDate >= " & vStartDate & " AND TranDate <= " & vEndDate & " order by trantype desc, trandate asc")
Dim Tax_Total, Taxable_Total, vGRS_Sales, vGrs_Tax_SL, vtotal_tax As Double
Tax_Total = 0: Taxable_Total = 0
If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
   rsORD_HIST.MoveFirst
   Do While Not rsORD_HIST.EOF
      vCustCode = Null2String(rsORD_HIST!custcode)
      vReference = Null2String(rsORD_HIST!trantype) & Null2String(rsORD_HIST!tranno)
      vRectype = "'D'"
      vTrantype = "'S'"
      vCust_Tin = "NULL"
      vCust_Name = Null2String(rsORD_HIST!custname)
      vGRS_Sales = N2Str2Zero(rsORD_HIST!TOTINVAMT)
      vTIN_Owner = TIN
      vRef_Name = Null2String(rsORD_HIST!custname)
      vTrandate = Null2String(rsORD_HIST!trandate)

      vGrs_Tax_SL = vGRS_Sales - N2Str2Zero(rsORD_HIST!DISCOUNT)
      'vTotal_Tax = N2Str2Zero(Ord_Hist!VAT)
      vtotal_tax = N2Str2Zero(rsORD_HIST!TOTINVAMT) / 1.1
      Set rsCustomer = New ADODB.Recordset
      Set rsCustomer = gconPMIOS.Execute("Select * from Customer Where CustCode = '" & vCustCode & "'")
      If Not rsCustomer.EOF And Not rsCustomer.BOF Then
         vAddress = Null2String(rsCustomer!custadrs)
         vCust_Addr1 = Left(vAddress, 30)
         vCust_Addr2 = Mid(vAddress, 31, 60)
         vLastName = "NULL"
         vFirstName = "NULL"
         vMidName = "NULL"
      Else
         vAddress = "NULL"
         vCust_Addr1 = "NULL"
         vCust_Addr2 = "NULL"
         vLastName = "NULL"
         vFirstName = "NULL"
         vMidName = "NULL"
      End If
  
      'vReference,vRectype,vTrantype,vCust_Tin,vCust_Name,vGRS_Sales,vTIN_Owner,vReg_Name,vTrandate,vCust_Addr1,vCust_Addr2,vLastName,vFirstName,vMidName
  
      Tax_Total = Tax_Total + vtotal_tax
      Taxable_Total = Taxable_Total + vGrs_Tax_SL

      gconBIRData.Execute "Insert Into BIR.DBF " & _
                          "(Reference,Rectype,Trantype,Cust_Tin,Cust_Name,GRS_Sales,TIN_Owner,Trandate,Cust_Addr1,Cust_Addr2)" & _
                          " Values ('" & vReference & "'," & vRectype & "," & vTrantype & "," & vCust_Tin & ",'" & vCust_Name & "'," & vGRS_Sales & "," & vTIN_Owner & ",'" & vTrandate & "'," & vCust_Addr1 & "," & vCust_Addr2 & ")"
      
      i = i + 1
      progCPB.Value = (i / rsORD_HIST.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsORD_HIST.MoveNext
   Loop
   labCPB.Caption = ""
   DoEvents
   Screen.MousePointer = 0
Else
   MsgSpeechBox "Error Opening Customer Order File"
   Exit Sub
End If

gconBIRData.Execute "Update BIR.DBF Set" & _
                  " Grs_Tax_SL = " & Taxable_Total & " ," & _
                  " Total_Tax = " & Tax_Total & _
                  " Where Reference = 'HEADER'"
End Sub

Private Sub cmdCheck_Click()
'Dim rsSales As ADODB.Recordset
'Set rsSales = New ADODB.Recordset
'Set rsSales = gconBIR_Relief.Execute("Select * from Sales")

Dim Vtax_month As String
Dim Vseq_no As Double
Dim Vtin As String
Dim Vregistered_name As String
Dim Vlast_name As String
Dim Vfirst_name As String
Dim Vmiddle_name As String
Dim Vaddress1 As String
Dim Vaddress2 As String
Dim Vgsales1 As Double
Dim Vgsales2 As Double
Dim Vgsales As Double
Dim Vgtsales As Double
Dim Vgesales As Double
Dim Vgzsales As Double
Dim Vtouttax As Double

Dim Vgpurchase As Double
Dim Vgtpurchase As Double
Dim Vgepurchase As Double
Dim Vgzpurchase As Double
Dim Vgtservpurchase As Double
Dim Vgtcappurchase As Double
Dim Vgtothpurchase As Double
Dim Vtinputtax As Double

Dim i As Integer

Dim vTINSeq_No As Integer
Vtax_month = ""
Vseq_no = 1
Vtin = ""
Vregistered_name = ""
Vlast_name = ""
Vfirst_name = ""
Vmiddle_name = ""
Vaddress1 = ""
Vaddress2 = ""
Vgsales = 0
Vgtsales = 0
Vgesales = 0
Vgzsales = 0
Vtouttax = 0

Vgpurchase = 0
Vgtpurchase = 0
Vgepurchase = 0
Vgzpurchase = 0
Vgtservpurchase = 0
Vgtcappurchase = 0
Vgtothpurchase = 0
Vtinputtax = 0

Dim LaborWSC As Double
Dim PartsMaterialsWSC As Double
If EXTRACT_TYPE = "SALES" Then
   Dim rsSales As ADODB.Recordset
   Set rsSales = New ADODB.Recordset
   Set rsSales = gconBIR_Relief.Execute("Select * from sales where tax_month = #" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "#")
   If Not rsSales.EOF And Not rsSales.BOF Then
      MsgBox "Critical Error: Existing Data of Parts & Service transaction already exist!" & vbCrLf & _
             "                Extraction is now Disabled."
      'gconBIR_Relief.Execute ("Delete * from Sales Where tax_month = '" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "'")
      cmdCheck.Enabled = False
      Exit Sub
   End If
   Set rsSales = New ADODB.Recordset
   Set rsSales = gconBIR_Relief.Execute("Select * from sales order by seq_no desc")
   If Not rsSales.EOF And Not rsSales.BOF Then
      Vseq_no = N2Str2Zero(rsSales!seq_no)
   Else
      Vseq_no = 1
   End If
   Dim rsORD_HIST As ADODB.Recordset
   Set rsORD_HIST = New ADODB.Recordset
   Set rsORD_HIST = gconPMIOS.Execute("Select * from ORD_HIST Where status = 'P' and (trantype = 'CSH' or trantype = 'CHG') and month(trandate) = " & What_month(cboMonth.Text) & " and year(trandate) = " & cboYear.Text)
   If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
      rsORD_HIST.MoveFirst
      Vseq_no = 0
      vTINSeq_No = 0
      i = 0
      Do While Not rsORD_HIST.EOF
         Vtax_month = N2Date2Null(lastDay(rsORD_HIST!trandate))
         Vseq_no = Vseq_no + 1
         vTINSeq_No = vTINSeq_No + 1
         'Vtin = "'1" & Format(Right(Year(rsORD_HIST!trandate), 2), "00") & Format(Month(rsORD_HIST!trandate), "000") & Format(vTINSeq_No, "000") & "'"
         Vtin = "''"
         Vregistered_name = N2Str2Null(rsORD_HIST!custname)
         Vlast_name = "''"
         Vfirst_name = "''"
         Vmiddle_name = "''"
         Vaddress1 = "'A'"
         Vaddress2 = "'B'"
         Vgsales = Round(N2Str2Zero(rsORD_HIST!netinvamt) - (Round(N2Str2Zero(rsORD_HIST!netinvamt) / 11, 2)), 2)
         Vgtsales = Round(N2Str2Zero(rsORD_HIST!netinvamt) - (Round(N2Str2Zero(rsORD_HIST!netinvamt) / 11, 2)), 2)
         Vgesales = 0
         Vgzsales = 0
         Vtouttax = Round((N2Str2Zero(rsORD_HIST!netinvamt) - (Round(N2Str2Zero(rsORD_HIST!netinvamt) / 11, 2))) * 0.1, 2)
         gconBIR_Relief.Execute "Insert into Sales " & _
                                "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gsales,gtsales,gesales,gzsales,touttax)" & _
                                " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgsales & "," & Vgtsales & "," & Vgesales & "," & Vgzsales & "," & Vtouttax & ")"
         i = i + 1
         progCPB.Value = (i / rsORD_HIST.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsORD_HIST.MoveNext
      Loop
   End If
   Dim rsRepor As ADODB.Recordset
   Dim rsRo_det As ADODB.Recordset
   Set rsRepor = New ADODB.Recordset
   Set rsRepor = gconCSMIOS.Execute("Select * from REPOR Where month(DTE_REL) = " & What_month(cboMonth.Text) & " and year(DTE_REL) = " & cboYear.Text)
   If Not rsRepor.EOF And Not rsRepor.BOF Then
      rsRepor.MoveFirst
      i = 0
      Do While Not rsRepor.EOF
         Vtax_month = N2Date2Null(lastDay(rsRepor!dte_rel))
         Vseq_no = Vseq_no + 1
         vTINSeq_No = vTINSeq_No + 1
         'Vtin = "'2" & Format(Right(Year(rsRepor!dte_rel), 2), "00") & Format(Month(rsRepor!dte_rel), "000") & Format(vTINSeq_No, "000") & "'"
         Vtin = "''"
         Vregistered_name = N2Str2Null(Left(rsRepor!Niym, 50))
         Vlast_name = "''"
         Vfirst_name = "''"
         Vmiddle_name = "''"
         Vaddress1 = "'A'"
         Vaddress2 = "'B'"
         Vgsales = Round(N2Str2Zero(rsRepor!ro_amount), 2)
         Set rsRo_det = New ADODB.Recordset
         Set rsRo_det = gconCSMIOS.Execute("Select SUM(DETPRC) - SUM(DISVAL) as TOTALChargeTOWSC from RO_DET Where livil = '1' and (WCODE = 'W' or WCODE = 'S' or WCODE = 'C') and rep_or = " & N2Str2Null(rsRepor!rep_or))
         If Not rsRo_det.EOF And Not rsRo_det.BOF Then
            LaborWSC = N2Str2Zero(rsRo_det!TOTALChargeTOWSC)
            Vgsales = Round(Vgsales + Round(LaborWSC, 2), 2)
         Else
            LaborWSC = 0
         End If
         Set rsRo_det = New ADODB.Recordset
         Set rsRo_det = gconCSMIOS.Execute("Select SUM(DETPRC*DETVOL)-SUM(DISVAL) as TOTALChargeTOWSC from RO_DET Where livil <> '1' and (WCODE = 'W' or WCODE = 'S' or WCODE = 'C') and rep_or = " & N2Str2Null(rsRepor!rep_or))
         If Not rsRo_det.EOF And Not rsRo_det.BOF Then
            PartsMaterialsWSC = N2Str2Zero(rsRo_det!TOTALChargeTOWSC)
            Vgsales = Round(Vgsales + Round(PartsMaterialsWSC, 2), 2)
         Else
            PartsMaterialsWSC = 0
         End If
         Vgsales1 = Round(Vgsales - (Round(Vgsales / 11, 2)), 2)
         Vgtsales = Round(Vgsales - (Round(Vgsales / 11, 2)), 2)
         Vtouttax = Round(Vgsales1 * 0.1, 2)
         Vgesales = 0
         Vgzsales = 0
         gconBIR_Relief.Execute "Insert into Sales " & _
                                "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gsales,gtsales,gesales,gzsales,touttax)" & _
                                " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgsales1 & "," & Vgtsales & "," & Vgesales & "," & Vgzsales & "," & Vtouttax & ")"
         i = i + 1
         progCPB.Value = (i / rsRepor.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsRepor.MoveNext
      Loop
   End If
End If

If EXTRACT_TYPE = "PURCHASE" Then
   Dim rsPurchase As ADODB.Recordset
   Set rsPurchase = New ADODB.Recordset
   Set rsPurchase = gconBIR_Relief.Execute("Select * from Purchase where tax_month = #" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "#")
   If Not rsPurchase.EOF And Not rsPurchase.BOF Then
      MsgBox "Critical Error: Existing Data of Purchases transaction already exist!" & vbCrLf & _
             "                Extraction is now Disabled."
      'gconBIR_Relief.Execute ("Delete * from Sales Where tax_month = '" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "'")
      cmdCheck.Enabled = False
      Exit Sub
   End If
   Set rsPurchase = New ADODB.Recordset
   Set rsPurchase = gconBIR_Relief.Execute("Select * from Purchase order by seq_no desc")
   If Not rsPurchase.EOF And Not rsPurchase.BOF Then
      Vseq_no = N2Str2Zero(rsPurchase!seq_no)
   Else
      Vseq_no = 1
   End If
   Dim rsJOURNAL_DET As ADODB.Recordset
   Dim rsJOURNAL_HD As ADODB.Recordset
   Dim rsVENDOR As ADODB.Recordset
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAMIS.Execute("Select * from JOURNAL_DET Where ACCT_CODE = '15-00001-00' AND JTYPE = 'APJ') and month(JDATE) = " & What_month(cboMonth.Text) & " and year(JDATE) = " & cboYear.Text)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.BOF Then
      rsJOURNAL_DET.MoveFirst
      vTINSeq_No = 0
      i = 0
      Do While Not rsJOURNAL_DET.EOF
         Vtax_month = N2Date2Null(lastDay(rsJOURNAL_DET!JDATE))
         Vseq_no = Vseq_no + 1
         Set rsJOURNAL_HD = New ADODB.Recordset
         Set rsJOURNAL_HD = gconAMIS.Execute("Select * from JOURNAL_HD Where JNo = " & N2Str2Null(rsJOURNAL_DET!JNo))
         If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.BOF Then
            Set rsVENDOR = New ADODB.Recordset
            Set rsVENDOR = gconAMIS.Execute("Select * from Vendor where code = " & N2Str2Null(rsJOURNAL_HD!VendorCode))
            If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
               Vtin = N2Str2Null(Left(Null2String(rsVENDOR!TIN), 9))
               Vregistered_name = N2Str2Null(rsVENDOR!nameofvendor)
               Vlast_name = N2Str2Null(rsVENDOR!lastname)
               Vfirst_name = N2Str2Null(rsVENDOR!firstname)
               Vmiddle_name = N2Str2Null(rsVENDOR!middlename)
               Vaddress1 = N2Str2Null(Left(Null2String(rsVENDOR!Address), 30))
               Vaddress2 = N2Str2Null(Mid(Null2String(rsVENDOR!Address), 31, 30))
            Else
               Vtin = "''"
               Vregistered_name = "''"
               Vlast_name = "''"
               Vfirst_name = "''"
               Vmiddle_name = "''"
               Vaddress1 = "''"
               Vaddress2 = "''"
            End If
         Else
            Vtin = "''"
            Vregistered_name = "''"
            Vlast_name = "''"
            Vfirst_name = "''"
            Vmiddle_name = "''"
            Vaddress1 = "''"
            Vaddress2 = "''"
         End If
               
         Vtinputtax = N2Str2Zero(rsJOURNAL_DET!Debit)
         Vgpurchase = N2Str2Zero(rsJOURNAL_DET!Debit) / 0.1
         Vgtpurchase = N2Str2Zero(rsJOURNAL_DET!Debit) / 0.1
         Vgepurchase = 0
         Vgzpurchase = 0
         Vgtservpurchase = 0
         Vgtcappurchase = 0
         Vgtothpurchase = N2Str2Zero(rsJOURNAL_DET!Debit) / 0.1
         
         'gpurchase,gtpurchase,gepurchase,gzpurchase,gtservpurchase,gtcappurchase,gtothpurchase,tinputtax
         gconBIR_Relief.Execute "Insert into Purchase " & _
                                "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gpurchase,gtpurchase,gepurchase,gzpurchase,gtservpurchase,gtcappurchase,gtothpurchase,tinputtax)" & _
                                " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgpurchase & "," & Vgtpurchase & "," & Vgepurchase & "," & Vgzpurchase & "," & Vgtservpurchase & "," & Vgtcappurchase & "," & Vgtothpurchase & "," & Vtinputtax & ")"
         i = i + 1
         progCPB.Value = (i / rsJOURNAL_DET.RecordCount) * 100
         labCPB.Caption = Int(progCPB.Value) & "% Completed"
         DoEvents
         rsJOURNAL_DET.MoveNext
      Loop
   End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrorCode
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
'txtFrom.Text = firstDay(LOGDATE)
'txtTo.Text = LOGDATE
fillcbomonth cboMonth
FillcboYear cboYear
cboMonth.Text = The_month(Month(LOGDATE))
cboYear.Text = Year(LOGDATE)
'FileCopy App.Path & "\bir.dbf", BIR_DATABASE_PATH & "BIR.dbf"
'Set gconBIRData = New ADODB.Connection
'gconBIRData.ConnectionString = BIRDATA_Connection
'frmSplash.labCon.Caption = "Connecting to BIR Database... Please wait..."
'DoEvents
'gconBIRData.Open
'Unload frmSplash
Set gconBIR_Relief = New ADODB.Connection
    gconBIR_Relief.ConnectionString = BIR_RELIEF_Connection
    frmSplash.labCon.Caption = "Connecting to BIR Database... Please wait..."
    DoEvents
    gconBIR_Relief.Open
    Unload frmSplash
Screen.MousePointer = 0
Exit Sub

ErrorCode:
Screen.MousePointer = 0
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
Unload frmSplash
cmdCheck.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
gconBIR_Relief.Close
End Sub
