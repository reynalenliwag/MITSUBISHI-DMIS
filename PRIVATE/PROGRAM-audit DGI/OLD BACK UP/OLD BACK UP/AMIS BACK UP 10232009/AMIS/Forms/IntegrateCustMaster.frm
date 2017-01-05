VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form frmIntegrateCustMaster 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integrate Customer Master File"
   ClientHeight    =   1530
   ClientLeft      =   1800
   ClientTop       =   1800
   ClientWidth     =   6090
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "IntegrateCustMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   765
      Left            =   5010
      Picture         =   "IntegrateCustMaster.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   945
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   765
      Left            =   4080
      Picture         =   "IntegrateCustMaster.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   690
      Width           =   945
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   330
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "IntegrateCustMaster.frx":0EDE
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "IntegrateCustMaster.frx":0EFA
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
      TabIndex        =   2
      Top             =   60
      Width           =   5835
   End
End
Attribute VB_Name = "frmIntegrateCustMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
Screen.MousePointer = 11
Dim i As Long
Dim CRIS_ACCOUNT_NO As String
Dim CRIS_CUSTCODE As String
Dim CRIS_COMPANY As String
Dim CRIS_CUST_NAME As String
Dim CRIS_LAST_NAME As String
Dim CRIS_FIRST_NAME As String
Dim CRIS_MIDDLE_NAME As String
Dim CRIS_SEX As String
Dim CRIS_ADDRESS As String
Dim CRIS_PROVINCIAL_ADD As String
Dim CRIS_ZIPCODE As String
Dim CRIS_PHONE As String
Dim CRIS_CELLNO As String
Dim CRIS_CATEGORY As String
Dim CRIS_PLATENO As String
Dim CRIS_OLDCODE As String

Dim rsCSMIOS_CUSTOMER As ADODB.Recordset
Dim rsPMIOS_CUSTOMER As ADODB.Recordset
Dim rsSMIS_CUSTOMER As ADODB.Recordset
Dim rsAMIS_CUSTOMER As ADODB.Recordset

Set rsCSMIOS_CUSTOMER = New ADODB.Recordset
Set rsCSMIOS_CUSTOMER = gconCSMIOS.Execute("Select * from CUSMAS Order by LastName,FirstName asc")
If Not rsCSMIOS_CUSTOMER.EOF And Not rsCSMIOS_CUSTOMER.BOF Then
   rsCSMIOS_CUSTOMER.MoveFirst: i = 0
   Do While Not rsCSMIOS_CUSTOMER.EOF
      CRIS_ACCOUNT_NO = "NULL"
      CRIS_CUSTCODE = N2Str2Null(SetCustomerCode(Null2String(rsCSMIOS_CUSTOMER!LASTNAME)))
      CRIS_COMPANY = N2Str2Null(rsCSMIOS_CUSTOMER!CUSCOMP)
      CRIS_CUST_NAME = N2Str2Null(rsCSMIOS_CUSTOMER!CUSNAM)
      CRIS_LAST_NAME = N2Str2Null(rsCSMIOS_CUSTOMER!LASTNAME)
      CRIS_FIRST_NAME = N2Str2Null(rsCSMIOS_CUSTOMER!FIRSTNAME)
      CRIS_MIDDLE_NAME = N2Str2Null(rsCSMIOS_CUSTOMER!MIDDLEINITIAL)
      CRIS_SEX = "NULL"
      CRIS_ADDRESS = N2Str2Null(rsCSMIOS_CUSTOMER!CUSADD)
      CRIS_PROVINCIAL_ADD = N2Str2Null(rsCSMIOS_CUSTOMER!PROVADD)
      CRIS_ZIPCODE = N2Str2Null(rsCSMIOS_CUSTOMER!CUSZIPC)
      CRIS_PHONE = N2Str2Null(rsCSMIOS_CUSTOMER!CUSPHON1)
      CRIS_CELLNO = "NULL"
      CRIS_CATEGORY = N2Str2Null(rsCSMIOS_CUSTOMER!CUSCAT)
      CRIS_PLATENO = N2Str2Null(rsCSMIOS_CUSTOMER!PLATENO)
      CRIS_OLDCODE = N2Str2Null(rsCSMIOS_CUSTOMER!CUSCDE)
      
      gconCRIS.Execute ("Insert Into CUSTOMER_MASTER " & _
                       "(ACCOUNT_NO,CUSTCODE,COMPANY,CUST_NAME,LAST_NAME,FIRST_NAME,MIDDLE_NAME,SEX,ADDRESS,PROVINCIAL_ADD,ZIPCODE,PHONE,CELLNO,CATEGORY,PLATENO,OLDCODE)" & _
                       " values (" & CRIS_ACCOUNT_NO & "," & CRIS_CUSTCODE & "," & CRIS_COMPANY & "," & CRIS_CUST_NAME & "," & CRIS_LAST_NAME & "," & CRIS_FIRST_NAME & "," & CRIS_MIDDLE_NAME & "," & CRIS_SEX & "," & CRIS_ADDRESS & "," & CRIS_PROVINCIAL_ADD & "," & CRIS_ZIPCODE & "," & CRIS_PHONE & "," & CRIS_CELLNO & "," & CRIS_CATEGORY & "," & CRIS_PLATENO & "," & CRIS_OLDCODE & ")")
      i = i + 1
      progCPB.Value = (i / rsCSMIOS_CUSTOMER.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsCSMIOS_CUSTOMER.MoveNext
   Loop
End If

Set rsSMIS_CUSTOMER = New ADODB.Recordset
Set rsSMIS_CUSTOMER = gconSMIS.Execute("Select * from CUSTOMER Order by LastName,FirstName asc")
If Not rsSMIS_CUSTOMER.EOF And Not rsSMIS_CUSTOMER.BOF Then
   rsSMIS_CUSTOMER.MoveFirst: i = 0
   Do While Not rsSMIS_CUSTOMER.EOF
      CRIS_ACCOUNT_NO = "NULL"
      CRIS_CUSTCODE = N2Str2Null(SetCustomerCode(Null2String(rsSMIS_CUSTOMER!LASTNAME)))
      CRIS_COMPANY = "NULL"
      CRIS_CUST_NAME = N2Str2Null(Null2String(rsSMIS_CUSTOMER!LASTNAME) & "," & Null2String(rsSMIS_CUSTOMER!FIRSTNAME) & " " & Null2String(rsSMIS_CUSTOMER!MIDDLEINITIAL))
      CRIS_LAST_NAME = N2Str2Null(rsSMIS_CUSTOMER!LASTNAME)
      CRIS_FIRST_NAME = N2Str2Null(rsSMIS_CUSTOMER!FIRSTNAME)
      CRIS_MIDDLE_NAME = N2Str2Null(rsSMIS_CUSTOMER!MIDDLEINITIAL)
      CRIS_SEX = N2Str2Null(rsSMIS_CUSTOMER!SEX)
      CRIS_ADDRESS = N2Str2Null(rsSMIS_CUSTOMER!CUSTOMERADD)
      CRIS_PROVINCIAL_ADD = N2Str2Null(rsSMIS_CUSTOMER!PROVINCIALADD)
      CRIS_ZIPCODE = N2Str2Null(rsSMIS_CUSTOMER!ZIPCODE)
      CRIS_PHONE = N2Str2Null(rsSMIS_CUSTOMER!TELEPHONENO)
      CRIS_CELLNO = "NULL"
      CRIS_CATEGORY = "NULL"
      CRIS_PLATENO = "NULL"
      CRIS_OLDCODE = N2Str2Null(rsSMIS_CUSTOMER!CODE)
      
      gconCRIS.Execute ("Insert Into CUSTOMER_MASTER " & _
                       "(ACCOUNT_NO,CUSTCODE,COMPANY,CUST_NAME,LAST_NAME,FIRST_NAME,MIDDLE_NAME,SEX,ADDRESS,PROVINCIAL_ADD,ZIPCODE,PHONE,CELLNO,CATEGORY,PLATENO,OLDCODE)" & _
                       " values (" & CRIS_ACCOUNT_NO & "," & CRIS_CUSTCODE & "," & CRIS_COMPANY & "," & CRIS_CUST_NAME & "," & CRIS_LAST_NAME & "," & CRIS_FIRST_NAME & "," & CRIS_MIDDLE_NAME & "," & CRIS_SEX & "," & CRIS_ADDRESS & "," & CRIS_PROVINCIAL_ADD & "," & CRIS_ZIPCODE & "," & CRIS_PHONE & "," & CRIS_CELLNO & "," & CRIS_CATEGORY & "," & CRIS_PLATENO & "," & CRIS_OLDCODE & ")")
      i = i + 1
      progCPB.Value = (i / rsSMIS_CUSTOMER.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsSMIS_CUSTOMER.MoveNext
   Loop
End If

Set rsAMIS_CUSTOMER = New ADODB.Recordset
Set rsAMIS_CUSTOMER = gconAmis.Execute("Select * from CUSTOMER Order by LastName,FirstName asc")
If Not rsAMIS_CUSTOMER.EOF And Not rsAMIS_CUSTOMER.BOF Then
   rsAMIS_CUSTOMER.MoveFirst: i = 0
   Do While Not rsAMIS_CUSTOMER.EOF
      CRIS_ACCOUNT_NO = N2Str2Null(rsAMIS_CUSTOMER!ACCOUNTNO)
      CRIS_CUSTCODE = N2Str2Null(SetCustomerCode(Null2String(rsAMIS_CUSTOMER!LASTNAME)))
      CRIS_COMPANY = "NULL"
      CRIS_CUST_NAME = N2Str2Null(rsAMIS_CUSTOMER!CUSTNAME)
      CRIS_LAST_NAME = N2Str2Null(rsAMIS_CUSTOMER!LASTNAME)
      CRIS_FIRST_NAME = N2Str2Null(rsAMIS_CUSTOMER!FIRSTNAME)
      CRIS_MIDDLE_NAME = N2Str2Null(rsAMIS_CUSTOMER!MIDDLENAME)
      CRIS_SEX = "NULL"
      CRIS_ADDRESS = N2Str2Null(rsAMIS_CUSTOMER!ADDRESS)
      CRIS_PROVINCIAL_ADD = N2Str2Null(rsAMIS_CUSTOMER!PROVINCIALADD)
      CRIS_ZIPCODE = "NULL"
      CRIS_PHONE = N2Str2Null(rsAMIS_CUSTOMER!PHONE)
      CRIS_CELLNO = N2Str2Null(rsAMIS_CUSTOMER!CELLNO)
      CRIS_CATEGORY = N2Str2Null(rsAMIS_CUSTOMER!CATEGORY)
      CRIS_PLATENO = N2Str2Null(rsAMIS_CUSTOMER!PLATENO)
      CRIS_OLDCODE = N2Str2Null(rsAMIS_CUSTOMER!CUSTCODE)
      
      gconCRIS.Execute ("Insert Into CUSTOMER_MASTER " & _
                       "(ACCOUNT_NO,CUSTCODE,COMPANY,CUST_NAME,LAST_NAME,FIRST_NAME,MIDDLE_NAME,SEX,ADDRESS,PROVINCIAL_ADD,ZIPCODE,PHONE,CELLNO,CATEGORY,PLATENO,OLDCODE)" & _
                       " values (" & CRIS_ACCOUNT_NO & "," & CRIS_CUSTCODE & "," & CRIS_COMPANY & "," & CRIS_CUST_NAME & "," & CRIS_LAST_NAME & "," & CRIS_FIRST_NAME & "," & CRIS_MIDDLE_NAME & "," & CRIS_SEX & "," & CRIS_ADDRESS & "," & CRIS_PROVINCIAL_ADD & "," & CRIS_ZIPCODE & "," & CRIS_PHONE & "," & CRIS_CELLNO & "," & CRIS_CATEGORY & "," & CRIS_PLATENO & "," & CRIS_OLDCODE & ")")
      i = i + 1
      progCPB.Value = (i / rsAMIS_CUSTOMER.RecordCount) * 100
      labCPB.Caption = Int(progCPB.Value) & "% Completed"
      DoEvents
      rsAMIS_CUSTOMER.MoveNext
   Loop
End If
Screen.MousePointer = 0
MsgBox "Finish!"
'=========================================================================================================
End Sub

Sub UpdateCustomerControl()
Dim NewCtlCde As String
Dim rsCUSTOMER_MASTER As ADODB.Recordset
Dim k As Integer
Screen.MousePointer = 11
gconCRIS.Execute "delete from CUSCTL"
For k = 65 To 90
    Set rsCUSTOMER_MASTER = New ADODB.Recordset
    Set rsCUSTOMER_MASTER = gconCRIS.Execute("Select CUSTCODE from CUSTOMER_MASTER where left(CUSTCODE,1) = '" & Chr(k) & "' order by CUSTCODE desc")
    If Not rsCUSTOMER_MASTER.EOF And Not rsCUSTOMER_MASTER.BOF Then
       NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCUSTOMER_MASTER!CUSTCODE, 2, 5)) + 1, "00000")
       gconCRIS.Execute "insert into cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
    Else
       gconCRIS.Execute "insert into cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
    End If
Next
Screen.MousePointer = 0
End Sub

Function SetCustomerCode(XXX As String)
UpdateCustomerControl
Dim rsCusCtl As ADODB.Recordset
Set rsCusCtl = New ADODB.Recordset
Set rsCusCtl = gconCRIS.Execute("select ctlcde from cusctl where left(ctlcde,1) = '" & Left(Trim(XXX), 1) & "'")
If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
   SetCustomerCode = Null2String(rsCusCtl!ctlcde)
End If
Set rsCusCtl = Nothing
End Function

Private Sub cmdExit_Click()
Unload Me
End Sub

Function SetAcctName(VVV As Variant) As String
Dim rsChartAccount2 As ADODB.Recordset
Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from ChartAccount where AcctCode = " & VVV, gconAmis, adOpenForwardOnly, adLockReadOnly
If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
   SetAcctName = UCase(Null2String(rsChartAccount2!Description))
Else
   SetAcctName = ""
End If
End Function

Function GetVoucherNo() As String
Dim rsJOURNAL_HD As ADODB.Recordset
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("Select * from Journal_HD Where Jtype = 'SJ' Order by VoucherNo desc")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.BOF Then
   GetVoucherNo = Format(NumericVal(rsJOURNAL_HD!VoucherNo) + 1, "000000")
End If
End Function

Private Sub Form_Load()
On Error GoTo ErrorCode
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
Screen.MousePointer = 0
Exit Sub

ErrorCode:
Screen.MousePointer = 0
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
Unload frmSplash
cmdCheck.Enabled = False
End Sub

