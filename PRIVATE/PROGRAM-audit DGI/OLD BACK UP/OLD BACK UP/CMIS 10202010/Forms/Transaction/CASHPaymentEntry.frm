VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "wizBox.ocx"
Object = "{205EA659-0BC9-4F44-85D9-FBC10C8940C1}#1.0#0"; "wizDigit.ocx"
Begin VB.Form frmCMISCASHPaymentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Official Receipt Payment Entry Box"
   ClientHeight    =   6195
   ClientLeft      =   75
   ClientTop       =   540
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CASHPaymentEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Payment Module"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   150
      TabIndex        =   13
      Top             =   5280
      Width           =   9015
      Begin VB.TextBox txtAmountTendered 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Tendered :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   14
         Top             =   390
         Width           =   2085
      End
   End
   Begin wizBox.Box Box1 
      Height          =   1695
      Left            =   30
      Top             =   60
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   9
      Top             =   3900
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit3 
         Height          =   1215
         Left            =   -420
         TabIndex        =   10
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   11
      Top             =   3900
      Width           =   8865
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   5
      Top             =   2160
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit2 
         Height          =   1215
         Left            =   -420
         TabIndex        =   6
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   7
      Top             =   2160
      Width           =   8865
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   420
      Width           =   9015
      Begin wizDigits.wizDigit wizDigit1 
         Height          =   1215
         Left            =   -420
         TabIndex        =   2
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9EFE3&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   300
      ScaleHeight     =   1245
      ScaleWidth      =   8865
      TabIndex        =   3
      Top             =   420
      Width           =   8865
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F88A56&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "P 2,030.00 "
         Top             =   90
         Width           =   1395
      End
   End
   Begin wizBox.Box Box2 
      Height          =   1695
      Left            =   30
      Top             =   1800
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin wizBox.Box Box3 
      Height          =   1695
      Left            =   30
      Top             =   3540
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2990
   End
   Begin VB.Label labAmount 
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   18
      Top             =   5670
      Width           =   465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   17
      Top             =   3600
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Tendered"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   1860
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipts Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   15
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "frmCMISCASHPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Save_Payment()
    If COMPANY_CODE = M_COMPANY_CODE Then
        Dim vReference                                      As String
        Dim PAID_NA                                         As Boolean
        If RECEIPTS_BALANCE <= 0 Then
            PAID_NA = 1
        Else
            PAID_NA = 0
        End If
        If MODE_OF_PAYMENT = "CASH" Then
            If TYPE_OF_PAYMENT = "FULL" Then
                SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                " CARDAMOUNT = 0," & _
                                " CHKAMOUNT = 0," & _
                                " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                " ReferenceNo = " & "NULL" & "," & _
                                " PAIDBY = " & "NULL" & "," & _
                                " CARDDATE = " & "NULL" & "," & _
                                " TOF1 = '1'," & _
                                " TOF2 = " & "NULL" & "," & _
                                " TOF3 = " & "NULL" & "," & _
                                " PAIDNA = 1" & _
                                " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT

                SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                                " PAIDNA = 1" & _
                                " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT

            ElseIf TYPE_OF_PAYMENT = "PARTIAL" Then
                If RECEIPTS_BALANCE <= 0 Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " CASHAMOUNT = ISNULL(CASHAMOUNT,0) + " & NumericVal(RECEIPTS_AMOUNT) - NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)) & "," & _
                                    " TOF1 = '1'," & _
                                    " PAIDNA = '" & PAID_NA & "' " & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT

                    SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                                    " PAIDNA = 1" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                Else
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " CASHAMOUNT = ISNULL(CASHAMOUNT,0) + " & NumericVal(RECEIPTS_AMOUNT) - NumericVal(RECEIPTS_BALANCE) & "," & _
                                    " TOF1 = '1'," & _
                                    " PAIDNA = '" & PAID_NA & "' " & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If
            End If
            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
            End If
            '========================================================

            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            End If
            '========================================================

            Call PostUnpostOR(OR_NUMBER_GLOBAL, "CASH", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)

        ElseIf MODE_OF_PAYMENT = "CHECK" Or MODE_OF_PAYMENT = "CARD" Then
            '============================
            If MODE_OF_PAYMENT = "CHECK" Then
                If TYPE_OF_PAYMENT = "FULL" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " CASHAMOUNT = 0," & _
                                    " CARDAMOUNT = 0," & _
                                    " CHKAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " ReferenceNo = " & "NULL" & "," & _
                                    " PAIDBY = " & "NULL" & "," & _
                                    " CARDDATE = " & "NULL" & "," & _
                                    " TOF1 = " & "NULL" & "," & _
                                    " TOF2 = '2'," & _
                                    " TOF3 = " & "NULL" & "," & _
                                    " PAIDNA = 1" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                ElseIf TYPE_OF_PAYMENT = "PARTIAL" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " CHKAMOUNT = ISNULL(CHKAMOUNT,0) + " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " TOF2 = '2'," & _
                                    " PAIDNA = '" & PAID_NA & "' " & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT

                    SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                                    " PAIDNA = '" & PAID_NA & "'" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If
                If vREFERENCENO <> "" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set BANK = " & N2Str2Null(CheckBankName(frmCMISOREntry.txtCUSCDE)) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If
                '========================================================
                'Updating Code:  JAA - 08272008
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                Else
                    NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                End If
                '========================================================

                Call PostUnpostOR(OR_NUMBER_GLOBAL, "CHECK", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)

            ElseIf MODE_OF_PAYMENT = "CARD" Then
                vReference = GetReferenceNo
                If TYPE_OF_PAYMENT = "FULL" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " CASHAMOUNT = 0," & _
                                    " CHKAMOUNT = 0," & _
                                    " CARDAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " TOF1 = " & "NULL" & "," & _
                                    " TOF2 = " & "NULL" & "," & _
                                    " TOF3 = '3'," & _
                                    " ReferenceNo = " & N2Str2Null(vReference) & "," & _
                                    " PAIDBY = 'N'," & _
                                    " BANK = " & N2Str2Null(xBankCode) & "," & _
                                    " PAIDNA = 1" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                ElseIf TYPE_OF_PAYMENT = "PARTIAL" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                    " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                    " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                    " CARDAMOUNT = ISNULL(CARDAMOUNT,0) + " & NumericVal(AMOUNT_TENDERED) & "," & _
                                    " TOF3 = '3'," & _
                                    " ReferenceNo = " & N2Str2Null(vReference) & "," & _
                                    " PAIDBY = 'N'," & _
                                    " BANK = " & N2Str2Null(CheckBankName(frmCMISOREntry.txtCUSCDE)) & "," & _
                                    " PAIDNA = '" & PAID_NA & "'" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT

                    SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                                    " PAIDNA = '" & PAID_NA & "'" & _
                                    " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If

                If vREFERENCENO = "" Then
                    SQL_STATEMENT = "update CMIS_Off_Dt Set ReferenceNo = " & N2Str2Null(vReference) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                Else
                    SQL_STATEMENT = "update CMIS_Off_Dt Set ReferenceNo = " & N2Str2Null(vREFERENCENO) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If
                '========================================================
                'Updating Code:  JAA - 08272008
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                Else
                    NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                End If
                '========================================================

                Call PostUnpostOR(OR_NUMBER_GLOBAL, "CARD", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)
            End If
            '=====================================

            '        SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                     '                        " CUTDATE = '" & CURRENT_CUST_CODE & "'," & _
                     '                        " VAT = " & VAT_OR & "," & _
                     '                        " PAIDNA = 1" & _
                     '                        " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
            '        gconDMIS.Execute SQL_STATEMENT

            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            End If
            '========================================================
        End If
    Else
        If MODE_OF_PAYMENT = "CASH" Then
            SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                            " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                            " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                            " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                            " CARDAMOUNT = 0," & _
                            " CHKAMOUNT = 0," & _
                            " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                            " TOF = '1'," & _
                            " VAT = " & VAT_OR & "," & _
                            " PAIDNA = 1" & _
                            " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
            gconDMIS.Execute SQL_STATEMENT

            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
            End If
            '========================================================

            SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                            " VAT = " & VAT_OR & "," & _
                            " PAIDNA = 1" & _
                            " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
            gconDMIS.Execute SQL_STATEMENT

            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            End If
            '========================================================

            Call PostUnpostOR(OR_NUMBER_GLOBAL, "CASH", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)

            '        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                     " CASH = ROUND(CASH,2) + " & RECEIPTS_AMOUNT & _
                     " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        ElseIf MODE_OF_PAYMENT = "CHECK" Or MODE_OF_PAYMENT = "CARD" Then
            If MODE_OF_PAYMENT = "CHECK" Then
                SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                " CASHAMOUNT = 0," & _
                                " CARDAMOUNT = 0," & _
                                " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                " TOF = '2'," & _
                                " VAT = " & VAT_OR & "," & _
                                " PAIDNA = 1" & _
                                " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT

                If vREFERENCENO <> "" Then
                    SQL_STATEMENT = "update CMIS_Off_Hd Set BANK = " & N2Str2Null(CheckBankName(frmCMISOREntry.txtCUSCDE)) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                    gconDMIS.Execute SQL_STATEMENT
                End If

                '========================================================
                'Updating Code:  JAA - 08272008
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                Else
                    NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                End If
                '========================================================

                Call PostUnpostOR(OR_NUMBER_GLOBAL, "CHECK", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)

                'gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                 " [CHECK] = ROUND([CHECK],2) + " & RECEIPTS_AMOUNT & _
                 " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf MODE_OF_PAYMENT = "CARD" Then
                vReference = GetReferenceNo
                SQL_STATEMENT = "update CMIS_Off_Hd Set" & _
                                " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                                " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                                " CASHAMOUNT = 0," & _
                                " CHKAMOUNT = 0," & _
                                " OR_AMT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                                " TOF = '3'," & _
                                " VAT = " & VAT_OR & "," & _
                                " ReferenceNo = " & N2Str2Null(vReference) & "," & _
                                " PAIDBY = 'N'," & _
                                " BANK = " & N2Str2Null(xBankCode) & "," & _
                                " PAIDNA = 1" & _
                                " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT

                '========================================================
                'Updating Code:  JAA - 08272008
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                Else
                    NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, ""
                End If
                '========================================================

                Call PostUnpostOR(OR_NUMBER_GLOBAL, "CARD", "POST", RECEIPTS_AMOUNT, TYPE_OF_PAYMENT)
                'gconDMIS.Execute ("update CMIS_Cash_Pos Set " & _
                 " CARD = ROUND(CARD,2) + " & RECEIPTS_AMOUNT & _
                 " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If

            SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                            " CUTDATE = '" & CURRENT_CUST_CODE & "'," & _
                            " VAT = " & VAT_OR & "," & _
                            " PAIDNA = 1" & _
                            " where VAT = " & VAT_OR & " AND OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
            gconDMIS.Execute SQL_STATEMENT

            If vREFERENCENO = "" Then
                SQL_STATEMENT = "update CMIS_Off_Dt Set ReferenceNo = " & N2Str2Null(vReference) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT
            Else
                SQL_STATEMENT = "update CMIS_Off_Dt Set ReferenceNo = " & N2Str2Null(vREFERENCENO) & " where OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
                gconDMIS.Execute SQL_STATEMENT
            End If

            '========================================================
            'Updating Code:  JAA - 08272008
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "P", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            Else
                NEW_LogAudit "P", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_HD"), "", Null2String(OR_NUMBER_GLOBAL), MODE_OF_PAYMENT, FindTransactionID(N2Str2Null(OR_NUMBER_GLOBAL), "OR_NUM", "CMIS_Off_DT")
            End If
            '========================================================
        End If
    End If

    MessagePop InfoFriend, "OR Information Updated", "OR Sucessfully Paid", 1000
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    If COMPANY_CODE = M_COMPANY_CODE Then
        CHANGE_DUE = 0
        FINAL_CASH = 0
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR))
        If MODE_OF_PAYMENT = "CASH" Then
            AMOUNT_TENDERED = 0
            wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
            wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
            txtAmountTendered.Text = NumericVal(RECEIPTS_AMOUNT) - CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)
        Else
            wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
            wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
            txtAmountTendered.Text = NumericVal(AMOUNT_TENDERED)
        End If
    Else
        AMOUNT_TENDERED = 0: CHANGE_DUE = 0
        wizDigit1.TextValue = ToDoubleNumber(RECEIPTS_AMOUNT)
        wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
        wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
        txtAmountTendered.Text = NumericVal(RECEIPTS_AMOUNT)
    End If
    If MODE_OF_PAYMENT = "CARD" Then txtAmountTendered.Locked = True
    If MODE_OF_PAYMENT = "CHECK" Then txtAmountTendered.Locked = True
End Sub

Private Sub txtAmountTendered_Change()
    RECEIPTS_AMOUNT = NumericVal(RECEIPTS_AMOUNT)
    AMOUNT_TENDERED = NumericVal(txtAmountTendered.Text)
    If AMOUNT_TENDERED >= RECEIPTS_AMOUNT Then
        CHANGE_DUE = AMOUNT_TENDERED - NumericVal(RECEIPTS_AMOUNT)
    Else
        CHANGE_DUE = "0.00"
    End If
    wizDigit2.TextValue = ToDoubleNumber(AMOUNT_TENDERED)
    wizDigit3.TextValue = ToDoubleNumber(CHANGE_DUE)
End Sub

Private Sub txtAmountTendered_GotFocus()
    If NumericVal(txtAmountTendered.Text) > 0 Then txtAmountTendered.Text = NumericVal(txtAmountTendered.Text) Else txtAmountTendered.Text = ""
End Sub

Private Sub txtAmountTendered_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn Then
        If COMPANY_CODE = M_COMPANY_CODE Then
            RECEIPTS_BALANCE = NumericVal(RECEIPTS_AMOUNT) - (NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR)) + NumericVal(AMOUNT_TENDERED))
            If NumericVal(AMOUNT_TENDERED) < NumericVal(RECEIPTS_AMOUNT) Then
                TYPE_OF_PAYMENT = "PARTIAL"
                If RECEIPTS_BALANCE <= 0 Then
                    If MsgBox("Figure Correct?", vbQuestion + vbYesNo, "CMIS") = vbYes Then
                        FINAL_CASH = NumericVal(RECEIPTS_AMOUNT) - NumericVal(CheckTotalPayment(OR_NUMBER_GLOBAL, VAT_OR))
                        Save_Payment
                        Unload Me
                        frmCMISOREntry.optCASH.Enabled = True
                        frmCMISOREntry.optCHECK.Enabled = True
                        frmCMISOREntry.optCARD.Enabled = True
                        frmCMISOREntry.RefreshDisplay
                        frmCMISOREntry.cmdPrint.Value = True
                    End If
                Else
                    If MsgBox("Payment Amount does not meet OR Amount..." & Chr(13) & "Continue?", vbExclamation + vbYesNo, "Message") = vbYes Then
                        FINAL_CASH = NumericVal(RECEIPTS_AMOUNT) - NumericVal(RECEIPTS_BALANCE)
                        Save_Payment
                        Unload Me
                        frmCMISOREntry.RefreshDisplay
                        frmCMISOREntry.picPayment.ZOrder 0
                        frmCMISOREntry.picPayment.Visible = True
                        frmCMISOREntry.optCASH.Enabled = False
                        frmCMISOREntry.optCHECK.Enabled = True
                        frmCMISOREntry.optCHECK.Value = True
                        frmCMISOREntry.optCHECK.SetFocus
                        frmCMISOREntry.optCARD.Value = False
                        frmCMISOREntry.optCANCEL.Value = False
                    Else
                        Exit Sub
                    End If
                End If
            Else
                If MsgBox("Figure Correct?", vbQuestion + vbYesNo, "CMIS") = vbYes Then
                    TYPE_OF_PAYMENT = "FULL"
                    AMOUNT_TENDERED = NumericVal(txtAmountTendered.Text)
                    RECEIPTS_AMOUNT = NumericVal(RECEIPTS_AMOUNT)
                    CHANGE_DUE = AMOUNT_TENDERED - NumericVal(RECEIPTS_AMOUNT)
                    Save_Payment
                    Unload Me
                    frmCMISOREntry.RefreshDisplay
                    frmCMISOREntry.cmdPrint.Value = True
                End If
            End If
        Else
            If AMOUNT_TENDERED < RECEIPTS_AMOUNT Then
                MsgBox "Payment Amount does not meet OR Amount...", vbInformation + vbOKOnly, "Message"
                Exit Sub
            End If
            If MsgBox("Figure Correct?", vbQuestion + vbYesNo, "CMIS") = vbYes Then
                AMOUNT_TENDERED = NumericVal(txtAmountTendered.Text)
                RECEIPTS_AMOUNT = NumericVal(RECEIPTS_AMOUNT)
                CHANGE_DUE = AMOUNT_TENDERED - NumericVal(RECEIPTS_AMOUNT)
                Save_Payment
                Unload Me
                frmCMISOREntry.RefreshDisplay
                frmCMISOREntry.cmdPrint.Value = True
            End If
        End If
    End If
End Sub

Private Sub txtAmountTendered_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAmountTendered_LostFocus()
    On Error Resume Next
    txtAmountTendered.SetFocus
    txtAmountTendered.Text = ToDoubleNumber(txtAmountTendered.Text)
End Sub

Function GetReferenceNo() As String
    Dim rsCMIS_OFF_HD                                       As ADODB.Recordset
    Set rsCMIS_OFF_HD = New ADODB.Recordset
    Set rsCMIS_OFF_HD = gconDMIS.Execute("Select CAST(ReferenceNo AS int) AS MAX_REFERENCENO from CMIS_Off_HD Order by MAX_REFERENCENO desc")
    If Not rsCMIS_OFF_HD.EOF And Not rsCMIS_OFF_HD.BOF Then
        GetReferenceNo = Format(NumericVal(rsCMIS_OFF_HD!MAX_REFERENCENO) + 1, "00000000")
    Else
        GetReferenceNo = "00000001"
    End If
End Function
