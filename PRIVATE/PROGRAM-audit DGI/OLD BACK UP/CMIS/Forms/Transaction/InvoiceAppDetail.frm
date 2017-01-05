VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInvoiceAppDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Application Detail"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12780
   ForeColor       =   &H00E0E0E0&
   Icon            =   "InvoiceAppDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   12735
      TabIndex        =   1
      Top             =   4350
      Width           =   12765
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
         Height          =   345
         Left            =   9630
         TabIndex        =   3
         Top             =   60
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
         Left            =   10830
         TabIndex        =   2
         Top             =   60
         Width           =   1785
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdInvoiceDetails 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   7488
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   -2147483633
      BackColorBkg    =   -2147483633
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

Function SetPartDesc(ILoveUMaam As Variant)
    Dim rsPartMas                                 As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("Select PartDesc from PMIS_partmas Where TYPE = 'P' and partno = '" & ILoveUMaam & "'")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartDesc = Null2String(rsPartMas!PartDesc)
    End If
End Function

Sub InitGrid()
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
    Dim LuvUMsChat                                As Long
    Dim MissUMsChat, luvNaLuvKita                 As Double
    Dim rsCSMS_HD                                 As ADODB.Recordset
    Dim rsDAYTRAN                                 As ADODB.Recordset
    If INVOICE_DETAIL_TYPE = "PI" Then
        Set rsDAYTRAN = New ADODB.Recordset
        Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'P' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'P' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        End If
        If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: InitGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                grdInvoiceDetails.AddItem "PARTS : " & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
        End If
    End If
    If INVOICE_DETAIL_TYPE = "AI" Then
        Set rsDAYTRAN = New ADODB.Recordset
        Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'A' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'A' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        End If
        If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: InitGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(NumericVal(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * (NumericVal(rsDAYTRAN!tranuprice) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
        End If
    End If
    If INVOICE_DETAIL_TYPE = "MI" Then
        Set rsDAYTRAN = New ADODB.Recordset
        Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_DAYTRAN Where TYPE = 'M' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        If rsDAYTRAN.EOF And rsDAYTRAN.BOF Then
            Set rsDAYTRAN = gconDMIS.Execute("Select * from PMIS_TDAYTRAN Where TYPE = 'M' AND Trantype = 'CSH' and tranno = '" & INVOICE_DETAIL_TRANNO & "' order by itemno asc")
        End If
        If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
            rsDAYTRAN.MoveFirst: InitGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
            Do While Not rsDAYTRAN.EOF
                LuvUMsChat = LuvUMsChat + 1
                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & " " & Null2String(rsDAYTRAN!stock_ord) & Chr(9) & " " & SetPartDesc(Null2String(rsDAYTRAN!stock_ord)) & Chr(9) & N2Str2Zero(rsDAYTRAN!tranqty) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!TRANUCOST)))
                MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!tranqty))
                luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!tranqty)) * NumericVal(N2Str2Zero(rsDAYTRAN!tranuprice) / 1.12))
                If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                rsDAYTRAN.MoveNext
            Loop
            labTotalQTY.Caption = MissUMsChat
            labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
        End If
    End If

    If INVOICE_DETAIL_TYPE = "SI" Then
        Set rsCSMS_HD = New ADODB.Recordset
        If INVOICE_DETAIL_TRANNO = "INT RO" Then

            Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM CSMS_REPOR where [TRANSTYPE] = 'R' AND REP_OR = '" & frmAMISJournalEntry_SJ.txtRefNo.Text & "'")

            'Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM CSMS_REPOR where [TRANSTYPE] = 'R'")

        Else
            Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM CSMS_REPOR where [TRANSTYPE] = 'R' AND INVOICE = '" & INVOICE_DETAIL_TRANNO & "'")
        End If
        If Not rsCSMS_HD.EOF And Not rsCSMS_HD.BOF Then
            Set rsDAYTRAN = New ADODB.Recordset
            Set rsDAYTRAN = gconDMIS.Execute("Select * from CSMS_RO_DET Where REP_OR = '" & Null2String(rsCSMS_HD!REP_OR) & "' order by livil asc, line_no asc")
            If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                rsDAYTRAN.MoveFirst: InitGrid: LuvUMsChat = 0: luvNaLuvKita = 0: MissUMsChat = 0
                Do While Not rsDAYTRAN.EOF
                    LuvUMsChat = LuvUMsChat + 1
                    If Null2String(rsDAYTRAN!livil) = "1" Then
                        If Null2Bool(rsCSMS_HD!VAT_EXEMPT) = True Then
                            grdInvoiceDetails.AddItem "LABOR:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detprc))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            luvNaLuvKita = luvNaLuvKita + N2Str2Zero(rsDAYTRAN!detprc)
                        Else
                            grdInvoiceDetails.AddItem "LABOR:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12))
                        End If
                    Else
                        If Null2Bool(rsCSMS_HD!VAT_EXEMPT) = True Then
                            If Null2String(rsDAYTRAN!livil) = "2" Then
                                grdInvoiceDetails.AddItem "PARTS:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                            If Null2String(rsDAYTRAN!livil) = "3" Then
                                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                            If Null2String(rsDAYTRAN!livil) = "4" Then
                                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                        Else
                            If Null2String(rsDAYTRAN!livil) = "2" Then
                                grdInvoiceDetails.AddItem "PARTS:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                            If Null2String(rsDAYTRAN!livil) = "3" Then
                                grdInvoiceDetails.AddItem "MATERIALS:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                            If Null2String(rsDAYTRAN!livil) = "4" Then
                                grdInvoiceDetails.AddItem "ACCESSORIES:" & Chr(9) & " " & Null2String(rsDAYTRAN!DETCDE) & Chr(9) & " " & Null2String(rsDAYTRAN!detdsc) & Chr(9) & N2Str2Zero(rsDAYTRAN!detvol) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsDAYTRAN!detprc) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detcost))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detcost)))
                            End If
                        End If
                        MissUMsChat = MissUMsChat + NumericVal(N2Str2Zero(rsDAYTRAN!detvol))
                        luvNaLuvKita = luvNaLuvKita + (NumericVal(N2Str2Zero(rsDAYTRAN!detvol)) * NumericVal(N2Str2Zero(rsDAYTRAN!detprc) / 1.12))
                    End If
                    If LuvUMsChat = 1 Then grdInvoiceDetails.RemoveItem 1
                    rsDAYTRAN.MoveNext
                Loop
                labTotalQTY.Caption = MissUMsChat
                labTotalAmount.Caption = ToDoubleNumber(luvNaLuvKita)
            End If
        End If
    End If

    If INVOICE_DETAIL_TYPE = "VI" Then
        Set rsCSMS_HD = New ADODB.Recordset
        Set rsCSMS_HD = gconDMIS.Execute("SELECT * FROM SMIS_PURCHAGREE where VI_NO = '" & INVOICE_DETAIL_TRANNO & "'")
        If Not rsCSMS_HD.EOF And Not rsCSMS_HD.BOF Then
            LuvUMsChat = LuvUMsChat + 1
            grdInvoiceDetails.AddItem "VEHICLES:" & Chr(9) & " " & Null2String(rsCSMS_HD!FRAMENO) & Chr(9) & " " & Null2String(rsCSMS_HD!Model) & Chr(9) & "1" & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMS_HD!NETSALESPRICE) / 1.12) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!NETSALESPRICE) / 1.12)) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!TOTAL_COST))) & Chr(9) & ToDoubleNumber(NumericVal(N2Str2Zero(rsCSMS_HD!TOTAL_COST)))
        End If
    End If

End Sub

