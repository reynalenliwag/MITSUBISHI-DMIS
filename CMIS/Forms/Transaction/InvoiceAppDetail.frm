VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInvoiceAppDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Application Detail - From Front-End"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13650
   ForeColor       =   &H00E0E0E0&
   Icon            =   "InvoiceAppDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   13650
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   4350
      Width           =   13605
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Quantity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1515
      End
      Begin VB.Label labTotalQTY 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11040
         TabIndex        =   3
         Top             =   40
         Width           =   1515
      End
      Begin VB.Label labTotalAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11760
         TabIndex        =   2
         Top             =   60
         Width           =   1785
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdInvoiceDetails 
      Height          =   4365
      Left            =   20
      TabIndex        =   0
      Top             =   0
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorSel    =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"InvoiceAppDetail.frx":09AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmInvoiceAppDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function SetPartDesc(ILoveUMaam As Variant, TranType As String)
    Dim rsPartMas                                           As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("SELECT STOCKDESC FROM PMIS_STOCKMAS WHERE TYPE='" & TranType & "' AND STOCKNO='" & ILoveUMaam & "'")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartDesc = Null2String(rsPartMas!STOCKDESC)
    End If
End Function

Sub initGrid()
    cleargrid grdInvoiceDetails
    If INVOICE_DETAIL_TYPE = "CSH" Then grdInvoiceDetails.FormatString = "  Part No.               |  Description                                    |  Qty     |  Unit Price         |  Total Amount   "
    If INVOICE_DETAIL_TYPE = "CHG" Then grdInvoiceDetails.FormatString = "  Part No.               |  Description                                    |  Qty     |  Unit Price         |  Total Amount   "
    If INVOICE_DETAIL_TYPE = "RO" Then grdInvoiceDetails.FormatString = "  Code No.               |  Description                                    |  Qty     |  Unit Price         |  Total Amount   "
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim LuvUMsChat                                          As Long
    Dim MissUMsChat, luvNaLuvKita                           As Double
    Dim rsCSMS_HD                                           As ADODB.Recordset
    Dim rsDAYTRAN                                           As ADODB.Recordset
    
    If INVOICE_DETAIL_TYPE = "PI" Then
        If INVOICE_DETAIL_TRANNO = "DR OUT" Then
            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'P' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'P' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "PARTS : " & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                grdInvoiceDetails.AddItem "PARTS : " & Chr(9) & _
                                          " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                          "-" & Chr(9) & _
                                          " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "P") & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        Else
            Set rsDAYTRAN = New ADODB.Recordset
                If COMPANY_CODE = "HCA" Then
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'P' AND Trantype IN ('CSH','CHG','RIV') and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                Else
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'P' AND Trantype IN ('CSH','CHG') and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                End If
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'P' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "PARTS : " & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                grdInvoiceDetails.AddItem "PARTS : " & Chr(9) & _
                                          " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                          "-" & Chr(9) & _
                                          " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "P") & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice)))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        End If
    End If
    If INVOICE_DETAIL_TYPE = "AI" Then
        If INVOICE_DETAIL_TRANNO = "DR OUT" Then
            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'A' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'A' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(NumericVal(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & _
                                          " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                          "-" & Chr(9) & _
                                          " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "A") & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        Else
            Set rsDAYTRAN = New ADODB.Recordset
                If COMPANY_CODE = "HCA" Then
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'A' AND Trantype IN ('CSH','CHG','RIV') and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                Else
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'A' AND Trantype IN ('CSH','CHG') and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                End If
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'A' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(NumericVal(rsDAYTRAN!tranuprice)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                 grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & _
                                            " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                            "-" & Chr(9) & _
                                            " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "A") & Chr(9) & _
                                            N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                            ToDoubleNumber(NumericVal(rsDAYTRAN!tranuprice)) & Chr(9) & _
                                            ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                            ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                            ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                                              
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice)))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        End If
    End If
    If INVOICE_DETAIL_TYPE = "MI" Then
        If INVOICE_DETAIL_TRANNO = "DR OUT" Then
            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'M' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
                Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'M' AND Trantype = 'DR' and tranno = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & _
                                          " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                          "-" & Chr(9) & _
                                          " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "M") & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                                                          
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        Else
            Set rsDAYTRAN = New ADODB.Recordset
                If COMPANY_CODE = "HCA" Then
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'M' AND Trantype IN ('CSH','CHG','RIV') and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                Else
                    Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'M' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
                End If
            If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
                Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'M' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
            End If
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                'OLD
                'grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & _
                                          " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & _
                                          "-" & Chr(9) & _
                                          " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord), "M") & Chr(9) & _
                                          N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & _
                                          ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice)))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            
            Total_AR_SLS_COST
            
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        End If
    End If

    If INVOICE_DETAIL_TYPE = "SI" Then
        Set rsCSMS_HD = New ADODB.Recordset
        If Left(INVOICE_DETAIL_TRANNO, 6) = "INT RO" Then
            Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM CSMS_REPOR where [TRANSTYPE] = 'R' AND REP_OR = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "'")
        Else
            Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM CSMS_REPOR where [TRANSTYPE] = 'R' AND INVOICE = '" & INVOICE_DETAIL_TRANNO & "'")
        End If
        If Not rsCSMS_HD.EOF And Not rsCSMS_HD.BOF Then
            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("Select * from CSMS_RO_DET Where REP_OR = '" & Null2String(rsCSMS_HD!REP_OR) & "' order by livil asc, line_no asc, jobtype asc")
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                rsDAYTRAN.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
                Do While Not rsDAYTRAN.EOF
                    LuvUMsChat = LuvUMsChat + 1
                    If Null2String(rsDAYTRAN!livil) = "1" Then
                        If Null2Bool(rsCSMS_HD!VAT_EXEMPT) = True Or Null2Bool(rsCSMS_HD!VAT_EXEMPT1) = True Then
                            'OLD
                            'grdInvoiceDetails.AddItem "LABOR:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & Chr(9) & ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DETPRC) / 1.12)), 2)) & Chr(9) & ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT) / 1.12)), 2)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            grdInvoiceDetails.AddItem "LABOR:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      "-" & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DETPRC) / 1.12)), 2)) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT) / 1.12)), 2)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))

                            luvNaLuvKita = luvNaLuvKita + N2Str2Zero(rsDAYTRAN!DETPRC)
                        Else
                            grdInvoiceDetails.AddItem "LABOR:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      "-" & Chr(9) & _
                                                      ToDoubleNumber(N2Str2Zero(rsDAYTRAN!DETPRC)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                                                      
                            luvNaLuvKita = luvNaLuvKita + NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))
                        End If
                    Else
                        If Null2Bool(rsCSMS_HD!VAT_EXEMPT) = True Or Null2Bool(rsCSMS_HD!VAT_EXEMPT1) = True Then
                            If Null2String(rsDAYTRAN!livil) = "2" Then
                                grdInvoiceDetails.AddItem "PARTS:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) _
                                                      & ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DETPRC)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                            
                            If Null2String(rsDAYTRAN!livil) = "3" Then
                                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DETPRC)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                            
                            If Null2String(rsDAYTRAN!livil) = "4" Then
                                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DETPRC)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(Round(NumericVal(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT)) / 1.12), 2)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                            
                        Else
                            If Null2String(rsDAYTRAN!livil) = "2" Then
                                grdInvoiceDetails.AddItem "PARTS:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) & _
                                                      ToDoubleNumber(N2Str2Zero(rsDAYTRAN!DETPRC)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                            
                            If Null2String(rsDAYTRAN!livil) = "3" Then
                                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) & _
                                                      ToDoubleNumber(N2Str2Zero(rsDAYTRAN!DETPRC)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                            
                            If Null2String(rsDAYTRAN!livil) = "4" Then
                                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!JOBTYPE) & Chr(9) & _
                                                      " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & _
                                                      N2Str2Zero(rsDAYTRAN!DETVOL) & Chr(9) & _
                                                      ToDoubleNumber(N2Str2Zero(rsDAYTRAN!DETPRC)) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST))) & Chr(9) & _
                                                      ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL)) * NumericVal(N2Str2Zero(rsDAYTRAN!DETCOST)))
                            End If
                        End If
                        MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!DETVOL))
                        luvNaLuvKita = luvNaLuvKita + NumericVal(N2Str2Zero(rsDAYTRAN!DET_AMT))
                    End If
                    If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                    rsDAYTRAN.MoveNext
                Loop
                
               Total_AR_SLS_COST
                
                labTotalQTY.Caption = MissUMsChat
                labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        End If
    End If

    If INVOICE_DETAIL_TYPE = "VI" Then
        Set rsCSMS_HD = New ADODB.Recordset
        Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM SMIS_PURCHAGREE where VI_NO = '" & INVOICE_DETAIL_TRANNO & "'")
        If Not rsCSMS_HD.EOF And Not rsCSMS_HD.BOF Then
            rsCSMS_HD.MoveFirst: initGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            LuvUMsChat = LuvUMsChat + 1
            grdInvoiceDetails.AddItem "VEHICLES:" & Chr(9) & _
                                " " & Null2String(rsCSMS_HD!FRAMENO) & Chr(9) & _
                                "-" & Chr(9) & _
                                " " & Null2String(rsCSMS_HD!Model) & Chr(9) & _
                                "1" & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMS_HD!NETSALESPRICE) / 1.12) & Chr(9) & _
                                ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!NETSALESPRICE) / 1.12)) & Chr(9) & _
                                ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!TOTAL_COST))) & Chr(9) & _
                                ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!TOTAL_COST)))
            If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
        End If
        
        'Total_AR_SLS_COST
    End If

End Sub

Sub Total_AR_SLS_COST()
    Dim SumAR As Double, SumSLS As Double, SumCOST As Double
    Dim ColAR As Long, ColSLS As Long, ColCOST As Long, lRow As Long
    
    ColAR = 5: ColSLS = 6: ColCOST = 8
    
    For lRow = grdInvoiceDetails.FixedRows To grdInvoiceDetails.Rows - 1
        SumAR = SumAR + CDbl(grdInvoiceDetails.TextMatrix(lRow, ColAR))
        SumSLS = SumSLS + CDbl(grdInvoiceDetails.TextMatrix(lRow, ColSLS))
        SumCOST = SumCOST + CDbl(grdInvoiceDetails.TextMatrix(lRow, ColCOST))
    Next lRow
        
    grdInvoiceDetails.AddItem Chr(9) & _
                              Chr(9) & _
                              Chr(9) & _
                              Chr(9) & _
                              Chr(9) & _
                              "" & ToDoubleNumber(SumAR) & "" & Chr(9) & _
                              "" & ToDoubleNumber(SumSLS) & "" & Chr(9) & _
                              Chr(9) & _
                              "" & ToDoubleNumber(SumCOST) & ""
                              
    'COLOR ONLY
    grdInvoiceDetails.Col = ColAR
        grdInvoiceDetails.Row = lRow
            grdInvoiceDetails.CellFontBold = True
                grdInvoiceDetails.CellForeColor = RGB(255, 0, 0)
                
    grdInvoiceDetails.Col = ColSLS
        grdInvoiceDetails.Row = lRow
            grdInvoiceDetails.CellFontBold = True
                grdInvoiceDetails.CellForeColor = RGB(255, 0, 0)
            
    grdInvoiceDetails.Col = ColCOST
        grdInvoiceDetails.Row = lRow
            grdInvoiceDetails.CellFontBold = True
                grdInvoiceDetails.CellForeColor = RGB(255, 0, 0)
End Sub
