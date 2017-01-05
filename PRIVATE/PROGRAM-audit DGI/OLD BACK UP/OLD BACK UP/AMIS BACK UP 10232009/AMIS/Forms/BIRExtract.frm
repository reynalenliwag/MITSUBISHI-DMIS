VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmBIRExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Data Extraction"
   ClientHeight    =   1530
   ClientLeft      =   210
   ClientTop       =   645
   ClientWidth     =   6090
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "BIRExtract.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6090
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
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
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   1110
      Width           =   1965
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   330
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "BIRExtract.frx":030A
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "BIRExtract.frx":0326
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5160
      MouseIcon       =   "BIRExtract.frx":0342
      MousePointer    =   99  'Custom
      Picture         =   "BIRExtract.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   690
      Width           =   720
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4455
      MouseIcon       =   "BIRExtract.frx":07FA
      MousePointer    =   99  'Custom
      Picture         =   "BIRExtract.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Process BIR Data Extraction"
      Top             =   690
      Width           =   720
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   2
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

'Upating Code       : AXP-0713200714:34
Private Sub cmdCheck_Click()
    If EXTRACT_TYPE = "SALES" Then
        If Function_Access(LOGID, "Acess_Process", "EXTRACT SALES ENTRIES TO BIR RELIEF") = False Then Exit Sub

    Else
        If Function_Access(LOGID, "Acess_Process", "EXTRACT PURCHASE ENTRIES TO BIR RELIEF") = False Then Exit Sub

    End If
    Dim Vtax_month                                                    As String
    Dim Vseq_no                                                       As Double
    Dim Vtin                                                          As String
    Dim Vregistered_name                                              As String
    Dim Vlast_name                                                    As String
    Dim Vfirst_name                                                   As String
    Dim Vmiddle_name                                                  As String
    Dim Vaddress1                                                     As String
    Dim Vaddress2                                                     As String
    Dim Vgsales                                                       As Double
    Dim Vgtsales                                                      As Double
    Dim Vgesales                                                      As Double
    Dim Vgzsales                                                      As Double
    Dim Vtouttax                                                      As Double

    Dim Vgpurchase                                                    As Double
    Dim Vgtpurchase                                                   As Double
    Dim Vgepurchase                                                   As Double
    Dim Vgzpurchase                                                   As Double
    Dim Vgtservpurchase                                               As Double
    Dim Vgtcappurchase                                                As Double
    Dim Vgtothpurchase                                                As Double
    Dim Vtinputtax                                                    As Double
    Dim vtax_rate                                                     As Double
    Dim i                                                             As Integer

    Dim VTOTAL_SALES, VTOTAL_DISCOUNT                                 As Double

    Dim vTINSeq_No                                                    As Integer
    On Error GoTo Errorcode:

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
    vtax_rate = 0

    Dim rsJournal_Det                                                 As ADODB.Recordset
    Dim rsSALES_LESS_DISCOUNT_JOURNAL_DET                             As ADODB.Recordset
    Dim rsJournal_HD                                                  As ADODB.Recordset
    Dim rsVENDOR                                                      As ADODB.Recordset
    Dim rsCUSTOMER                                                    As ADODB.Recordset

    If EXTRACT_TYPE = "SALES" Then
        Dim rsSales                                                   As ADODB.Recordset
        Set rsSales = New ADODB.Recordset
        Set rsSales = gconBIR_RELIEF.Execute("Select * from sales where tax_month = #" & lastDay(STR(What_month(cboMonth.Text)) & "/1/" & cboYear.Text) & "#")
        If Not rsSales.EOF And Not rsSales.BOF Then
            MsgBox "Critical Error: Existing Data of Parts && Service transaction already exist!" & vbCrLf & _
                 "                Extraction is now Disabled."
            'gconBIR_Relief.Execute ("delete from Sales Where tax_month = '" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "'")
            cmdCheck.Enabled = False
            Exit Sub
        End If
        Set rsSales = New ADODB.Recordset
        Set rsSales = gconBIR_RELIEF.Execute("Select * from sales order by seq_no desc")
        If Not rsSales.EOF And Not rsSales.BOF Then
            Vseq_no = N2Str2Zero(rsSales!seq_no)
        Else
            Vseq_no = 1
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("Select JNO,JDATE from AMIS_Journal_Det Where status = 'P' and (LEFT(ACCT_CODE,2) = '41' OR LEFT(ACCT_CODE,2) = '42') AND JTYPE = 'SJ' and month(JDATE) = " & What_month(cboMonth.Text) & " and year(JDATE) = " & cboYear.Text & " GROUP BY JNO,JDATE")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            rsJournal_Det.MoveFirst
            Vseq_no = 0
            vTINSeq_No = 0
            i = 0
            Do While Not rsJournal_Det.EOF
                Vtax_month = N2Date2Null(lastDay(rsJournal_Det!Jdate))
                Vseq_no = Vseq_no + 1
                vTINSeq_No = vTINSeq_No + 1
                Set rsJournal_HD = New ADODB.Recordset
                Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where JNo = " & N2Str2Null(rsJournal_Det!JNo))
                If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                    Set rsCUSTOMER = New ADODB.Recordset
                    Set rsCUSTOMER = gconDMIS.Execute("Select * from ALL_CUSTMASTER_AMIS where CUSTCODE = " & N2Str2Null(rsJournal_HD!CustomerCode))
                    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
                        'If Null2String(rsCUSTOMER!TIN) <> "" Then
                        '   Vtin = N2Str2Null(Left(Null2String(rsCUSTOMER!TIN), 9))
                        'Else
                        'Vtin = Format(Right(Year(rsJOURNAL_DET!jdate), 3), "000") & Format(Month(rsJOURNAL_DET!jdate), "000") & Format(vTINSeq_No, "000")
                        Vtin = "''"
                        'End If
                        'Vregistered_name = N2Str2Null(Left(Null2String(rsCustomer!AcctName), 50))
                        Vregistered_name = "''"
                        If Null2String(rsCUSTOMER!lastname) <> "" Then
                            Vlast_name = N2Str2Null(Left(Null2String(rsCUSTOMER!lastname), 30))
                        Else
                            Vlast_name = "''"
                        End If
                        If Null2String(rsCUSTOMER!Firstname) <> "" Then
                            Vfirst_name = N2Str2Null(Left(Null2String(rsCUSTOMER!Firstname), 30))
                        Else
                            Vfirst_name = "''"
                        End If
                        If Null2String(rsCUSTOMER!MiddleName) <> "" Then
                            Vmiddle_name = N2Str2Null(Left(Null2String(rsCUSTOMER!MiddleName), 30))
                        Else
                            Vmiddle_name = "''"
                        End If
                        Vaddress1 = "'A'"
                        Vaddress2 = "'B'"
                        'If Len(Null2String(rsCustomer!ADDRESS)) > 30 Then
                        '   Vaddress1 = N2Str2Null(Left(Null2String(rsCustomer!ADDRESS), 30))
                        '   Vaddress2 = N2Str2Null(Mid(Null2String(rsCustomer!ADDRESS), 31, 30))
                        'Else
                        '   If Null2String(rsCustomer!ADDRESS) <> "" Then
                        '      Vaddress1 = N2Str2Null(rsCustomer!ADDRESS)
                        '   Else
                        '      Vaddress1 = "''"
                        '   End If
                        '   Vaddress2 = "''"
                        'End If
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
                Set rsSALES_LESS_DISCOUNT_JOURNAL_DET = New ADODB.Recordset
                Set rsSALES_LESS_DISCOUNT_JOURNAL_DET = gconDMIS.Execute("Select SUM(CREDIT) AS TOTAL_SALES from AMIS_Journal_Det where (left(acct_code,2) = '41' or left(acct_code,2) = '42') AND Jno = '" & rsJournal_Det!JNo & "'")
                If Not rsSALES_LESS_DISCOUNT_JOURNAL_DET.EOF And Not rsSALES_LESS_DISCOUNT_JOURNAL_DET.BOF Then
                    VTOTAL_SALES = N2Str2Zero(rsSALES_LESS_DISCOUNT_JOURNAL_DET!TOTAL_SALES)
                End If
                Set rsSALES_LESS_DISCOUNT_JOURNAL_DET = New ADODB.Recordset
                Set rsSALES_LESS_DISCOUNT_JOURNAL_DET = gconDMIS.Execute("Select SUM(DEBIT) AS TOTAL_DISCOUNT from AMIS_Journal_Det where (left(acct_code,2) = '51' or left(acct_code,2) = '52') AND Jno = '" & rsJournal_Det!JNo & "'")
                If Not rsSALES_LESS_DISCOUNT_JOURNAL_DET.EOF And Not rsSALES_LESS_DISCOUNT_JOURNAL_DET.BOF Then
                    VTOTAL_DISCOUNT = N2Str2Zero(rsSALES_LESS_DISCOUNT_JOURNAL_DET!TOTAL_DISCOUNT)
                End If
                Set rsSALES_LESS_DISCOUNT_JOURNAL_DET = Nothing
                Vgsales = Round(VTOTAL_SALES - VTOTAL_DISCOUNT, 2)
                Vgtsales = Round(VTOTAL_SALES - VTOTAL_DISCOUNT, 2)
                Vgesales = 0
                Vgzsales = 0
                Vtouttax = Round((VTOTAL_SALES - VTOTAL_DISCOUNT) * VatPercentRate(VAT_RATE), 2)
                vtax_rate = VatPercentRate(VAT_RATE)
                gconBIR_RELIEF.Execute "Insert into Sales " & _
                                       "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gsales,gtsales,gesales,gzsales,touttax,tax_rate)" & _
                                     " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgsales & "," & Vgtsales & "," & Vgesales & "," & Vgzsales & "," & Vtouttax & "," & vtax_rate & ")"
                i = i + 1
                progCPB.Value = (i / rsJournal_Det.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsJournal_Det.MoveNext
            Loop
        End If
        MsgBox "Extraction Completed!", vbInformation, "Done"
    End If
    If EXTRACT_TYPE = "PURCHASE" Then
        Dim rsPurchase                                                As ADODB.Recordset
        Set rsPurchase = New ADODB.Recordset
        Set rsPurchase = gconBIR_RELIEF.Execute("Select * from Purchase where tax_month = #" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "#")
        If Not rsPurchase.EOF And Not rsPurchase.BOF Then
            MsgBox "Critical Error: Existing Data of Purchases transaction already exist!" & vbCrLf & _
                 "                Extraction is now Disabled."
            'gconBIR_Relief.Execute ("delete from Sales Where tax_month = '" & lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text)) & "'")
            cmdCheck.Enabled = False
            Exit Sub
        End If
        'Set rsPurchase = New ADODB.Recordset
        'Set rsPurchase = gconBIR_Relief.Execute("Select * from Purchase order by seq_no desc")
        'If Not rsPurchase.EOF And Not rsPurchase.BOF Then
        '   Vseq_no = N2Str2Zero(rsPurchase!seq_no)
        'Else
        '   Vseq_no = 1
        'End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det Where status = 'P' and LEFT(ACCT_CODE,5) = '11-05' AND JTYPE = 'APJ' and month(JDATE) = " & What_month(cboMonth.Text) & " and year(JDATE) = " & cboYear.Text)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            rsJournal_Det.MoveFirst
            Vseq_no = 0
            vTINSeq_No = 0
            i = 0
            Do While Not rsJournal_Det.EOF
                Vtax_month = N2Date2Null(lastDay(rsJournal_Det!Jdate))
                Vseq_no = Vseq_no + 1
                vTINSeq_No = vTINSeq_No + 1
                Set rsJournal_HD = New ADODB.Recordset
                Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where JNo = " & N2Str2Null(rsJournal_Det!JNo))
                If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                    Set rsVENDOR = New ADODB.Recordset
                    Set rsVENDOR = gconDMIS.Execute("Select * from ALL_Vendor where code = " & N2Str2Null(rsJournal_HD!VendorCode))
                    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
                        If Null2String(rsVENDOR!TIN) <> "" Then
                            Vtin = N2Str2Null(Left(Null2String(rsVENDOR!TIN), 9))
                        Else
                            'Vtin = Format(Right(Year(rsJOURNAL_DET!jdate), 3), "000") & Format(Month(rsJOURNAL_DET!jdate), "000") & Format(vTINSeq_No, "000")
                            Vtin = "''"
                        End If
                        Vregistered_name = N2Str2Null(rsVENDOR!nameofvendor)
                        If Null2String(rsVENDOR!lastname) <> "" Then
                            Vlast_name = N2Str2Null(rsVENDOR!lastname)
                        Else
                            Vlast_name = "''"
                        End If
                        If Null2String(rsVENDOR!Firstname) <> "" Then
                            Vfirst_name = N2Str2Null(rsVENDOR!Firstname)
                        Else
                            Vfirst_name = "''"
                        End If
                        If Null2String(rsVENDOR!MiddleName) <> "" Then
                            Vmiddle_name = N2Str2Null(rsVENDOR!MiddleName)
                        Else
                            Vmiddle_name = "''"
                        End If
                        If Len(Null2String(rsVENDOR!Address)) > 30 Then
                            Vaddress1 = N2Str2Null(Left(Null2String(rsVENDOR!Address), 30))
                            Vaddress2 = N2Str2Null(Mid(Null2String(rsVENDOR!Address), 31, 30))
                        Else
                            If Null2String(rsVENDOR!Address) <> "" Then
                                Vaddress1 = N2Str2Null(rsVENDOR!Address)
                            Else
                                Vaddress1 = "''"
                            End If
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
                Else
                    Vtin = "''"
                    Vregistered_name = "''"
                    Vlast_name = "''"
                    Vfirst_name = "''"
                    Vmiddle_name = "''"
                    Vaddress1 = "''"
                    Vaddress2 = "''"
                End If
                Vtinputtax = Round(N2Str2Zero(rsJournal_Det!DEBIT), 2)
                Vgpurchase = Round(N2Str2Zero(rsJournal_Det!DEBIT) / VatPercentRate(VAT_RATE), 2)
                Vgtpurchase = Round(N2Str2Zero(rsJournal_Det!DEBIT) / VatPercentRate(VAT_RATE), 2)
                Vgepurchase = 0
                Vgzpurchase = 0
                Vgtservpurchase = 0
                Vgtcappurchase = 0
                Vgtothpurchase = Round(N2Str2Zero(rsJournal_Det!DEBIT) / VatPercentRate(VAT_RATE), 2)
                vtax_rate = VatPercentRate(VAT_RATE)
                'gpurchase,gtpurchase,gepurchase,gzpurchase,gtservpurchase,gtcappurchase,gtothpurchase,tinputtax
                gconBIR_RELIEF.Execute "Insert into Purchase " & _
                                       "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gpurchase,gtpurchase,gepurchase,gzpurchase,gtservpurchase,gtcappurchase,gtothpurchase,tinputtax,tax_rate)" & _
                                     " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgpurchase & "," & Vgtpurchase & "," & Vgepurchase & "," & Vgzpurchase & "," & Vgtservpurchase & "," & Vgtcappurchase & "," & Vgtothpurchase & "," & Vtinputtax & "," & vtax_rate & ")"
                i = i + 1
                progCPB.Value = (i / rsJournal_Det.RecordCount) * 100
                labCPB.Caption = Int(progCPB.Value) & "% Completed"
                DoEvents
                rsJournal_Det.MoveNext
            Loop
        End If
        MsgBox "Extraction Completed!", vbInformation, "Done"
    End If
    LogAudit "R", "B.I.R. DATA EXTRACTION", cboMonth & "-" & cboYear
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Sub UselessSabiNiLaisol()
'   Dim rsORD_HIST As ADODB.Recordset
'   Set rsORD_HIST = New ADODB.Recordset
'   Set rsORD_HIST = gconDMIS.Execute("Select * from ORD_HIST Where (trantype = 'CSH' or trantype = 'CHG') and month(trandate) = " & What_month(cboMonth.Text) & " and year(trandate) = " & cboYear.Text)
'   If Not rsORD_HIST.EOF And Not rsORD_HIST.BOF Then
'      rsORD_HIST.MoveFirst
'      vTINSeq_No = 0
'      i = 0
'      Do While Not rsORD_HIST.EOF
'         Vtax_month = N2Date2Null(lastDay(rsORD_HIST!TRANDATE))
'         Vseq_no = Vseq_no + 1
'         vTINSeq_No = vTINSeq_No + 1
'         Vtin = "''"
'         'Vtin = Right(Year(rsORD_HIST!TRANDATE), 3) & Format(Month(rsORD_HIST!TRANDATE), "000") & Format(vTINSeq_No, "000")
'         'Vtin = "'005001002'"
'         Vregistered_name = N2Str2Null(Left(Null2String(rsORD_HIST!AcctName), 50))
'         Vlast_name = "''"
'         Vfirst_name = "''"
'         Vmiddle_name = "''"
'         Vaddress1 = "'A'"
'         Vaddress2 = "'B'"
'         Vgsales = Round(N2Str2Zero(rsORD_HIST!NETINVAMT) - (N2Str2Zero(rsORD_HIST!NETINVAMT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE))), 2)
'         Vgtsales = Round(N2Str2Zero(rsORD_HIST!NETINVAMT) - (N2Str2Zero(rsORD_HIST!NETINVAMT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE))), 2)
'         Vgesales = 0
'         Vgzsales = 0
'         Vtouttax = Round((N2Str2Zero(rsORD_HIST!NETINVAMT) - (N2Str2Zero(rsORD_HIST!NETINVAMT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE)))) * VatPercentRate(VAT_RATE), 2)
'         gconBIR_Relief.Execute "Insert into Sales " & _
          '                                "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gsales,gtsales,gesales,gzsales,touttax,tax_rate)" & _
          '                                " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgsales & "," & Vgtsales & "," & Vgesales & "," & Vgzsales & "," & Vtouttax & "," & 0 & ")"
'         i = i + 1
'         progCPB.Value = (i / rsORD_HIST.RecordCount) * 100
'         labCPB.Caption = Int(progCPB.Value) & "% Completed"
'         DoEvents
'         rsORD_HIST.MoveNext
'      Loop
'   End If
'   Dim rsRepor As ADODB.Recordset
'   Set rsRepor = New ADODB.Recordset
'   Set rsRepor = gconDMIS.Execute("Select * from REPOR Where month(DTE_REL) = " & What_month(cboMonth.Text) & " and year(DTE_REL) = " & cboYear.Text)
'   If Not rsRepor.EOF And Not rsRepor.BOF Then
'      rsRepor.MoveFirst
'      vTINSeq_No = 0
'      i = 0
'      Do While Not rsRepor.EOF
'         Vtax_month = N2Date2Null(lastDay(rsRepor!DTE_REL))
'         Vseq_no = Vseq_no + 1
'         vTINSeq_No = vTINSeq_No + 1
'         Vtin = "''"
'         'Vtin = Right(Year(rsRepor!DTE_REL), 3) & Format(Month(rsRepor!DTE_REL), "000") & Format(vTINSeq_No, "000")
'         'Vtin = "'005001002'"
'         Vregistered_name = N2Str2Null(Left(Null2String(rsRepor!NIYM), 6))
'         Vlast_name = "''"
'         Vfirst_name = "''"
'         Vmiddle_name = "''"
'         Vaddress1 = "'A'"
'         Vaddress2 = "'B'"
'         Vgsales = Round(N2Str2Zero(rsRepor!RO_AMOUNT) - (N2Str2Zero(rsRepor!RO_AMOUNT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE))), 2)
'         Vgtsales = Round(N2Str2Zero(rsRepor!RO_AMOUNT) - (N2Str2Zero(rsRepor!RO_AMOUNT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE))), 2)
'         Vgesales = 0
'         Vgzsales = 0
'         Vtouttax = Round((N2Str2Zero(rsRepor!RO_AMOUNT) - (N2Str2Zero(rsRepor!RO_AMOUNT) / (ConvertToBIRDecimalFormat(VAT_RATE) / VatPercentRate(VAT_RATE)))) * VatPercentRate(VAT_RATE), 2)
'         vtax_rate = VAT_RATE
'         gconBIR_Relief.Execute "Insert into Sales " & _
          '                                "(tax_month,seq_no,tin,registered_name,last_name,first_name,middle_name,address1,address2,gsales,gtsales,gesales,gzsales,touttax,tax_rate)" & _
          '                                " values (" & Vtax_month & "," & Vseq_no & "," & Vtin & "," & Vregistered_name & "," & Vlast_name & "," & Vfirst_name & "," & Vmiddle_name & "," & Vaddress1 & "," & Vaddress2 & "," & Vgsales & "," & Vgtsales & "," & Vgesales & "," & Vgzsales & "," & Vtouttax & "," & 0 & ")"
'         i = i + 1
'         progCPB.Value = (i / rsRepor.RecordCount) * 100
'         labCPB.Caption = Int(progCPB.Value) & "% Completed"
'         DoEvents
'         rsRepor.MoveNext
'      Loop
'   End If
'End Sub

Private Sub Form_Load()
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
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
    Set gconBIR_RELIEF = New ADODB.Connection
    gconBIR_RELIEF.ConnectionString = BIR_RELIEF_Connection
    frmSplash.labCon.Caption = "Connecting to BIR Database... Please wait..."
    DoEvents
    gconBIR_RELIEF.Open
    Unload frmSplash
    Screen.MousePointer = 0

    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gconBIR_RELIEF.Close
End Sub

