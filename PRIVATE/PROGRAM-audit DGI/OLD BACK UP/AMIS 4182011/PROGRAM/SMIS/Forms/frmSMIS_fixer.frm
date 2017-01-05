VERSION 5.00
Begin VB.Form frmSMIS_Fixer 
   Caption         =   "SMIS FIXER"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFixer 
      Caption         =   "SMIS FIXER"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmSMIS_fixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSalesOrder        As ADODB.Recordset
Dim rsMRRINV            As ADODB.Recordset
Dim SO_ignkey_no        As String
Dim SO_invoiceddate     As String
Dim SO_releaseddate     As String
Dim SO_vi_no            As String
Dim SO_customer_code    As String
Dim SO_prospectID       As String

Private Sub cmdFixer_Click()
     If MsgBox("Are You Sure You Want To Update Master File ", vbInformation + vbYesNo) = vbYes Then
            Set rsSalesOrder = gconDMIS.Execute("Select ignkey_no,invoiceddate,datereleased,vi_no,code,ProspectID from SMIS_SALESORDER where datereleased is not null")
                
            Do While Not rsSalesOrder.EOF
                SO_ignkey_no = N2Str2Null(rsSalesOrder!IGNKEY_NO)
                SO_invoiceddate = N2Date2Null(rsSalesOrder!InvoicedDate)
                SO_releaseddate = N2Date2Null(rsSalesOrder!DateReleased)
                SO_vi_no = N2Str2Null(rsSalesOrder!vi_no)
                SO_customer_code = N2Str2Null(rsSalesOrder!CODE)
                SO_prospectID = N2Str2Null(rsSalesOrder!PROSPECTID)
                
                    Set rsMRRINV = gconDMIS.Execute("Select ignkey  from SMIS_MrrInv_table where ignkey = " & SO_ignkey_no & "")
                    
                    gconDMIS.Execute ("Update SMIS_SalesOrder set STATUS = 'P', SOSTATUS = 'P' where ignkey_no = " & SO_ignkey_no & " and STATUS <> 'C'")
                    gconDMIS.Execute ("Update SMIS_Mrrinv_table set invoiceddate = " & SO_invoiceddate & ",datereleased = " & SO_releaseddate & ",vi_no =" & SO_vi_no & ", CustomerCode = " & SO_customer_code & ",ProspectID = " & SO_prospectID & ", released = 1, ISTATUS = 'R'  where ignkey = " & N2Str2Null(rsMRRINV!ignkey) & " ")
                    gconDMIS.Execute ("Update CRIS_PROSPECTS Set invoiceno = " & SO_vi_no & ", logclosingdate = " & SO_releaseddate & ", Status = 'C' where PROSPECTID = " & SO_prospectID & "")
                    
                    rsSalesOrder.MoveNext
                    Me.Caption = ">>" & SO_ignkey_no
            Loop
            
            MsgBox "AUTOMATIC YAN...", vbOKOnly, "SMIS"
            Unload Me
            Set rsSalesOrder = Nothing
            Set rsMRRINV = Nothing
     Else
        Exit Sub
     End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub
