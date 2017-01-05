VERSION 5.00
Begin VB.Form frmCMIS_viewInvoiceDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Invoiced Customer Detail"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCMIS_viewInvoiceCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3900
      TabIndex        =   7
      Top             =   1110
      Width           =   1305
   End
   Begin VB.TextBox txt_Name 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   975
      TabIndex        =   5
      Top             =   1710
      Width           =   4200
   End
   Begin VB.TextBox txt_Code 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   1710
      Width           =   765
   End
   Begin VB.TextBox txt_InvNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   2
      Top             =   1110
      Width           =   1770
   End
   Begin VB.ComboBox cboInvoiceType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   690
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter invoice number from Parts, Accessories, Materials, Vehicle and Service to view Customer Detail"
      Height          =   405
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type Invoice Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1140
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Select Invoice Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   690
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   30
      Top             =   30
      Width           =   5235
   End
End
Attribute VB_Name = "frmCMIS_viewInvoiceDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xENTITYCODE                                                  As String
Private Sub cboInvoiceType_Click()
    txt_Code = ""
    txt_Name = ""
End Sub

Private Sub cboInvoiceType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboInvoiceType.Text = "" Then
            MsgBox "Please select Invoice Type...", vbInformation, "System Messege"
            Exit Sub
        Else
            txt_InvNo.SetFocus
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Errorcode:
    txt_InvNo = Format(txt_InvNo, "000000")
    txt_Code = ""
    txt_Name = ""

    Dim TTYPE                                                    As String
    Dim INVTYPE                                                  As String
    Dim RSCUSTOMERINFO                                           As ADODB.Recordset
    Dim RSACCOUNTNAME                                            As ADODB.Recordset
    Select Case cboInvoiceType.ListIndex
    Case 0
        TTYPE = "P": INVTYPE = "CSH"
        TRANTYPE = "PARTS INVOICE"
    Case 1
        TTYPE = "P": INVTYPE = "CHG"
        TRANTYPE = "PARTS INVOICE"
    Case 2
        TTYPE = "P": INVTYPE = "DR"
        TRANTYPE = "PARTS INVOICE"
    Case 3
        TTYPE = "M": INVTYPE = "CSH"
        TRANTYPE = "MATERIALS INVOICE"
    Case 4
        TTYPE = "M": INVTYPE = "CHG"
        TRANTYPE = "MATERIALS INVOICE"
    Case 5
        TTYPE = "M": INVTYPE = "DR"
        TRANTYPE = "MATERIALS INVOICE"
    Case 6
        TTYPE = "A": INVTYPE = "CSH"
        TRANTYPE = "ACCESSORIES INVOICE"
    Case 7
        TTYPE = "A": INVTYPE = "CHG"
        TRANTYPE = "ACCESSORIES INVOICE"
    Case 8
        TTYPE = "A": INVTYPE = "DR"
        TRANTYPE = "ACCESSORIES INVOICE"
    Case 9
        TTYPE = "": INVTYPE = "VI"
        TRANTYPE = "VEHICLE INVOICE"
    Case 10
        TTYPE = "": INVTYPE = "SI"
        TRANTYPE = "SERVICE INVOICE"
    Case 11
        TTYPE = "": INVTYPE = "UI"
        TRANTYPE = "USER DEFINE INVOICE"
    End Select

    If INVTYPE = "VI" Then
        Set RSCUSTOMERINFO = gconDMIS.Execute("SELECT CODE FROM SMIS_SALESORDER WHERE VI_NO=" & N2Str2Null(Format(txt_InvNo, "000000")))
    ElseIf INVTYPE = "UI" Then
        Set RSCUSTOMERINFO = gconDMIS.Execute("SELECT CUSTOMERCODE AS CODE,ENTITY_CLASS FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND STATUS = 'P' AND INVOICENO=" & N2Str2Null(Format(txt_InvNo, "000000000")))
    ElseIf INVTYPE = "SI" Then
        Set RSCUSTOMERINFO = gconDMIS.Execute("SELECT ACCT_NO AS CODE  FROM CSMS_REPOR WHERE INVOICE=" & N2Str2Null(Format(txt_InvNo, "000000")))
    Else
        Set RSCUSTOMERINFO = gconDMIS.Execute("SELECT CUSTCODE CODE FROM PMIS_VW_ISS_HISTORY WHERE STATUS='P' AND TYPE=" & N2Str2Null(TTYPE) & " AND TRANTYPE=" & N2Str2Null(INVTYPE) & " AND TRANNO=" & N2Str2Null(Format(txt_InvNo, "000000")))
    End If

    If Not (RSCUSTOMERINFO.EOF Or RSCUSTOMERINFO.BOF) Then
        txt_Code = Null2String(RSCUSTOMERINFO!code)
        
'        If COMPANY_CODE = "HMH" Then
'            xENTITYCODE = Null2String(RSCUSTOMERINFO!ENTITY_CLASS)
'            Set RSACCOUNTNAME = gconDMIS.Execute("SELECT ACCOUNTNAME AS ACCTNAME FROM ALL_ENTITY WHERE CODE=" & N2Str2Null(txt_Code) & " AND ENTITYCODE =" & N2Str2Null(xENTITYCODE))
'        Else
            Set RSACCOUNTNAME = gconDMIS.Execute("SELECT ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(txt_Code))
'        End If
        If Not (RSACCOUNTNAME.EOF Or RSACCOUNTNAME.BOF) Then
            txt_Name = Null2String(RSACCOUNTNAME!AcctName)
            Clipboard.SetText (txt_Name)
            frmCMISOREntry.txtCUSCDE.Text = txt_Code
            frmCMISOREntry.cboCUSNAME.Text = txt_Name
            frmCMISOREntry.cmdSave.Value = True
            frmCMISOREntry.txtReference.Text = txt_InvNo
            Unload Me
            frmCMISOREntry.txtReference.SetFocus
        End If
    Else
        MessagePop InfoWarning, "Invoice No. not found.", "Please check Invoice Type / Invoice No."
    End If
    Exit Sub
Errorcode:
    MsgBox Err.Description
    Err.Clear

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    With cboInvoiceType
        .AddItem "PARTS INVOICE-CASH"
        .AddItem "PARTS INVOICE-CHARGE"
        .AddItem "PARTS INVOICE-DR"

        .AddItem "MATERIALS INVOICE-CASH"
        .AddItem "MATERIALS INVOICE-CHARGE"
        .AddItem "MATERIALS INVOICE-DR"

        .AddItem "ACCESSORIES INVOICE-CASH"
        .AddItem "ACCESSORIES INVOICE-CHARGE"
        .AddItem "ACCESSORIES INVOICE-DR"

        .AddItem "VEHICLE INVOICE"
        .AddItem "SERVICE INVOICE"
        .AddItem "USER DEFINE INVOICE"
        .ListIndex = 0
    End With
End Sub

Private Sub txt_Code_Change()

End Sub

Private Sub TXT_INVNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub
