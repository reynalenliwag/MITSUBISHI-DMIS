VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAMIS_Payment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Journal payment details"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   Icon            =   "frmAMIS_Payment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdlAllledger 
      Height          =   5055
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorSel    =   16711680
      BackColorBkg    =   14737632
      SelectionMode   =   1
      AllowUserResizing=   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8010
      TabIndex        =   2
      Top             =   5160
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5730
      TabIndex        =   1
      Top             =   5190
      Width           =   2175
   End
End
Attribute VB_Name = "frmAMIS_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub initAllLedger()
     With grdlAllledger
        If JOURNALTYPE = "SJ" Then
            .ColWidth(0) = 1200
            .ColWidth(1) = 1500
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
            .ColWidth(4) = 1200
            .ColWidth(5) = 1200
            .ColWidth(6) = 2500
            .Row = 0
            .Col = 0: .Text = "Journal Date"
            .Col = 1: .Text = "Voucher No"
            .Col = 2: .Text = "Type"
            .Col = 3: .Text = "Invoice Date"
            .Col = 4: .Text = "Invoice Type"
            .Col = 5: .Text = "Invoice No"
            .Col = 6: .Text = "Amount"
        ElseIf JOURNALTYPE = "APJ" Then
            .ColWidth(0) = 1200
            .ColWidth(1) = 1500
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
            .ColWidth(4) = 1200
            .ColWidth(5) = 1200
            .ColWidth(6) = 2500
            .Row = 0
            .Col = 0: .Text = "Due Date"
            .Col = 1: .Text = "CDJ Voucher No"
            .Col = 2: .Text = "N/A"
            .Col = 3: .Text = "N/A"
            .Col = 4: .Text = "Doc date"
            .Col = 5: .Text = "AP voucherno"
            .Col = 6: .Text = "Amount"
         End If
    End With
End Sub
Private Sub Form_Load()
 CenterMe frmMain, Me, 1
 initAllLedger
End Sub
Sub FillPaymentdetail(xInvoiceNo As String, xinvoiceType As String)
    Dim rsdetail As New ADODB.Recordset
    Dim nard As Integer
    Dim TotalAmount As Double
cleargrid:        initAllLedger
    If JOURNALTYPE = "SJ" Then
        Set rsdetail = gconDMIS.Execute("select * from AMIS_CRJ_detail where invoiceno='" & xInvoiceNo & "' and invoicetype ='" & xinvoiceType & "'")
      ElseIf JOURNALTYPE = "APJ" Then
        Set rsdetail = gconDMIS.Execute("select * from AMIS_CV_detail where pv_voucherno='" & xInvoiceNo & "'")
    End If
    
    nard = 0
    TotalAmount = 0
    If Not (rsdetail.EOF And rsdetail.BOF) Then
        Do While Not rsdetail.EOF
                nard = nard + 1
                If JOURNALTYPE = "SJ" Then
                grdlAllledger.AddItem (rsdetail!jdate) & Chr(9) & (rsdetail!voucherno) & Chr(9) & _
                                      (rsdetail!InvoiceType) & Chr(9) & (rsdetail!invoicedate) & Chr(9) & _
                                      (rsdetail!InvoiceType) & Chr(9) & (rsdetail!invoiceno) & Chr(9) & _
                                      (ToDoubleNumber(rsdetail!invoiceamount))
                Else
                grdlAllledger.AddItem (rsdetail!duedate) & Chr(9) & (rsdetail!voucherno) & Chr(9) & _
                                      "" & Chr(9) & "" & Chr(9) & _
                                      (rsdetail!docdate) & Chr(9) & (rsdetail!pv_voucherno) & Chr(9) & _
                                      (ToDoubleNumber(rsdetail!amount))
                End If
                If JOURNALTYPE = "SJ" Then
                    TotalAmount = TotalAmount + (NumericVal(rsdetail!invoiceamount))
                Else
                    TotalAmount = TotalAmount + (NumericVal(rsdetail!amount))
                End If
            rsdetail.MoveNext
        Loop
         If nard > 0 Then grdlAllledger.RemoveItem 1
         lbltotal.Caption = ToDoubleNumber(TotalAmount)
    End If
    Set rsdetail = Nothing
End Sub

