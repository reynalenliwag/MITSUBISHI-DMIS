VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISLEDGERVendors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendors Ledger"
   ClientHeight    =   7950
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11850
   ForeColor       =   &H00FFFFFF&
   Icon            =   "VendorsLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11850
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2700
      TabIndex        =   3
      Top             =   0
      Width           =   9105
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   9
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox txtCode3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2850
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtCode2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtCode1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtNameofVendor 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1620
         MaxLength       =   35
         TabIndex        =   14
         Top             =   570
         Width           =   7380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2670
         TabIndex        =   12
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2070
         TabIndex        =   11
         Top             =   240
         Width           =   135
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   10
         Top             =   210
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   180
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5955
      Left            =   2670
      TabIndex        =   15
      Top             =   960
      Width           =   9135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   90
         ScaleHeight     =   495
         ScaleWidth      =   8895
         TabIndex        =   22
         Top             =   5340
         Width           =   8895
         Begin VB.TextBox txtTotalBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   7020
            MaxLength       =   20
            TabIndex        =   25
            Top             =   60
            Width           =   1785
         End
         Begin VB.TextBox txtTotalDebit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   4260
            MaxLength       =   20
            TabIndex        =   24
            Top             =   60
            Width           =   1395
         End
         Begin VB.TextBox txtTotalCredit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   23
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Height          =   255
            Left            =   3150
            TabIndex        =   26
            Top             =   90
            Width           =   1395
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdAccountsLedger 
         Height          =   5085
         Left            =   60
         TabIndex        =   16
         Top             =   180
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   8969
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   16744448
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "VendorsLedger.frx":030A
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7890
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2565
      Begin Crystal.CrystalReport rptVendor 
         Left            =   1500
         Top             =   6360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox TextSearch 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   1
         Top             =   180
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstVendor 
         Height          =   7155
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   12621
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "VendorsLedger.frx":0624
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "VENDOR NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
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
      Left            =   11010
      MouseIcon       =   "VendorsLedger.frx":0786
      MousePointer    =   99  'Custom
      Picture         =   "VendorsLedger.frx":08D8
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit Window"
      Top             =   7020
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   10320
      MouseIcon       =   "VendorsLedger.frx":0C3E
      MousePointer    =   99  'Custom
      Picture         =   "VendorsLedger.frx":0D90
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Print this Record"
      Top             =   7035
      Width           =   705
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
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
      Left            =   9630
      MouseIcon       =   "VendorsLedger.frx":10F6
      MousePointer    =   99  'Custom
      Picture         =   "VendorsLedger.frx":1248
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Find a Record"
      Top             =   7035
      Width           =   705
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
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
      Left            =   8940
      MouseIcon       =   "VendorsLedger.frx":1542
      MousePointer    =   99  'Custom
      Picture         =   "VendorsLedger.frx":1694
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Move to Next Record"
      Top             =   7035
      Width           =   705
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Prev"
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
      Left            =   8250
      MouseIcon       =   "VendorsLedger.frx":19EC
      MousePointer    =   99  'Custom
      Picture         =   "VendorsLedger.frx":1B3E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Move to Previous Record"
      Top             =   7035
      Width           =   705
   End
End
Attribute VB_Name = "frmAMISLEDGERVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVENDOR                                                          As ADODB.Recordset
Dim rsJournal_HDDet                                                   As ADODB.Recordset
Dim AddorEdit, ORDER_BY                                               As String
Attribute ORDER_BY.VB_VarUserMemId = 1073938434
Dim TUTAL_DEBIT, TUTAL_CREDIT, TUTAL_BALANCE                          As Double
Attribute TUTAL_DEBIT.VB_VarUserMemId = 1073938436
Attribute TUTAL_CREDIT.VB_VarUserMemId = 1073938436
Attribute TUTAL_BALANCE.VB_VarUserMemId = 1073938436

Function SetVendorName(VVV As Variant)
    Dim rsVendorDup                                                   As ADODB.Recordset
    Set rsVendorDup = New ADODB.Recordset
    rsVendorDup.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVendorDup.EOF And Not rsVendorDup.BOF Then SetVendorName = Null2String(rsVendorDup!Nameofvendor) Else SetVendorName = ""
End Function

Sub FillGrids()
    Dim OUTBALANCE                                                    As Double
    Dim Reference                                                     As String
    Dim cnt                                                           As Integer
    Dim cnt_adjusment As Integer
    Dim tmp_voucher As String
    cleargrid grdAccountsLedger: InitGrid
    TUTAL_BALANCE = 0: TUTAL_BALANCE = 0: cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0
    cnt_adjusment = 0

    Set rsJournal_HDDet = New ADODB.Recordset
        rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_Det.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.jtype  where (AMIS_Journal_Hd.Status = 'P' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '21-01' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02') OR ((AMIS_Journal_HD.JTYPE = 'VCJ' AND AMIS_Journal_HD.Debit = 0) OR (AMIS_Journal_HD.JTYPE = 'VDJ' AND AMIS_Journal_HD.Credit = 0)))) AND AMIS_Journal_Hd.VendorCode = '" & txtcode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        rsJournal_HDDet.MoveFirst
        Do While Not rsJournal_HDDet.EOF
            cnt = cnt + 1
            
            
            
            If Null2String(rsJournal_HDDet!jtype) = "VDJ" Or Null2String(rsJournal_HDDet!jtype) = "VCJ" Then
                 tmp_voucher = Null2String(rsJournal_HDDet!VOUCHERNO)
                 If tmp_voucher = Null2String(rsJournal_HDDet!VOUCHERNO) Then
                     cnt_adjusment = cnt_adjusment + 1
                 End If
            End If
            Reference = Null2String(rsJournal_HDDet!jtype) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
            If Null2String(rsJournal_HDDet!jtype) = "VPJ" Then
                OUTBALANCE = OUTBALANCE + N2Str2Zero(rsJournal_HDDet!amounttopay)
            Else
                If Null2String(rsJournal_HDDet!jtype) = "VCJ" Then
                     If N2Str2Zero(rsJournal_HDDet!DEBIT) = 0 And cnt_adjusment >= 1 Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + N2Str2Zero(rsJournal_HDDet!CM)
                     End If
                ElseIf Null2String(rsJournal_HDDet!jtype) = "VDJ" Then
                    If N2Str2Zero(rsJournal_HDDet!DEBIT) = 0 And cnt_adjusment >= 1 Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) - N2Str2Zero(rsJournal_HDDet!DM)
                    End If
                Else
                    OUTBALANCE = (OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!CREDIT) - N2Str2Zero(rsJournal_HDDet!DEBIT)))
                End If
            End If
           
            If Null2String(rsJournal_HDDet!jtype) = "VCJ" Then
                If N2Str2Zero(rsJournal_HDDet!DEBIT) = 0 And cnt_adjusment >= 1 Then
                    grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                             Reference & Chr(9) & _
                                            " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                              "0.00" & Chr(9) & _
                                              ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CM)) & Chr(9) & _
                                              ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID
          
                cnt_adjusment = 0
                End If
              ElseIf Null2String(rsJournal_HDDet!jtype) = "VDJ" Then
                      If N2Str2Zero(rsJournal_HDDet!DEBIT) = 0 And cnt_adjusment >= 1 Then
                        grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                        " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DM)) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID
                     cnt_adjusment = 0
                     End If
             Else
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                        " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT)) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT)) & Chr(9) & _
                                          ToDoubleNumber(N2Str2IntZero(OUTBALANCE)) & Chr(9) & rsJournal_HDDet!ID
            End If
            If Null2String(rsJournal_HDDet!jtype) = "VCJ" Then
                TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CM)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "VDJ" Then
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DM)
            Else
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
            End If
            rsJournal_HDDet.MoveNext
        Loop

        If cnt > 0 Then grdAccountsLedger.RemoveItem 1

    Else
        cleargrid grdAccountsLedger
    End If
    txtTotalDebit.Text = ToDoubleNumber(TUTAL_DEBIT)
    txtTotalCredit.Text = ToDoubleNumber(TUTAL_CREDIT)
    txtTotalBalance.Text = ToDoubleNumber(TUTAL_BALANCE + OUTBALANCE)
    
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsVendors                                                     As ADODB.Recordset
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendors = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    'Set rsVendors = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD INNER JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE (dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ') AND dbo.ALL_Vendor.nameofvendor like '" & XXX & "%' ORDER BY dbo.ALL_Vendor.NameofVendor")
    Set rsVendors = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD left outer JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE (dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ' OR dbo.AMIS_Journal_HD.JType = 'VPJ' ) AND dbo.ALL_Vendor.nameofvendor like '" & XXX & "%' ORDER BY dbo.ALL_Vendor.NameofVendor")
    If Not (rsVendors.EOF And rsVendors.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendors
        lstVendor.Refresh
        lstVendor.Enabled = True
    Else
        lstVendor.Enabled = False
    End If
End Sub

Sub InitGrid()
    With grdAccountsLedger
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1300: .ColWidth(2) = 1400
        .ColWidth(3) = 1400: .ColWidth(4) = 1400: .ColWidth(5) = 1800
        .ColWidth(6) = 1: .Row = 0
        .Col = 0: .Text = "DOCDATE"
        .Col = 1: .Text = "REFERENCE"
        .Col = 2: .Text = "CHECKNO"
        .Col = 3: .Text = "DEBIT"
        .Col = 4: .Text = "CREDIT"
        .Col = 5: .Text = "BALANCE"
        .Col = 6: .Text = "ID"
    End With
End Sub

Sub InitMemVars()
    Frame1.Enabled = True
    txtcode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
    txtNameofVendor.Text = "":
    txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
    txtTotalBalance.Text = ZERO:
End Sub

Sub rsRefresh()
    Set rsVENDOR = New ADODB.Recordset
    'rsVENDOR.Open "SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD INNER JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ' ORDER BY dbo.ALL_Vendor.NameofVendor", gconDMIS, adOpenKeyset
    rsVENDOR.Open "SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD left outer JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ' OR dbo.AMIS_Journal_HD.JType = 'VPJ' ORDER BY dbo.ALL_Vendor.NameofVendor", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemvars()
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        Frame1.Enabled = False
        labID.Caption = Null2String(rsVENDOR!ID)
        txtcode.Text = Null2String(rsVENDOR!code)
        txtCode1.Text = Mid(Null2String(rsVENDOR!code), 1, 3)
        txtCode2.Text = Mid(Null2String(rsVENDOR!code), 5, 2)
        txtCode3.Text = Mid(Null2String(rsVENDOR!code), 8, 3)
        txtNameofVendor.Text = Null2String(rsVENDOR!Nameofvendor)
        FillGrids
    End If
End Sub

Sub PrintToExcel()
    Dim OUTBALANCE                                                    As Double
    Dim Reference                                                     As String
    Dim cnt                                                           As Integer
    Dim xlApp
    Dim xlBook
    Dim xlSheet1

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "VendorledgerFile.xls")
    Set xlSheet1 = xlBook.Worksheets(1)



    'cleargrid grdAccountsLedger: InitGrid
    TUTAL_BALANCE = 0: TUTAL_BALANCE = 0: cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0


    Set rsJournal_HDDet = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        rsJournal_HDDet.Open "select AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_Det.ID,AMIS_Journal_Det.JNo,AMIS_Journal_Det.JDate,AMIS_Journal_Det.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Det.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_Det inner Join AMIS_Journal_Hd on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo where AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Hd.VendorCode = '" & txtcode.Text & "' order by AMIS_Journal_Det.jdate asc,AMIS_Journal_Det.id asc", gconDMIS
    Else
        rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_Det.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo where (AMIS_Journal_Hd.Status = 'P' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '21-01' OR Left(AMIS_Journal_Det.Acct_Code,5) = '21-02')) OR (AMIS_Journal_HD.JTYPE = 'VCJ' OR AMIS_Journal_HD.JTYPE = 'VDJ')) AND AMIS_Journal_Hd.VendorCode = '" & txtcode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
    End If
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        rsJournal_HDDet.MoveFirst
        Do While Not rsJournal_HDDet.EOF
            cnt = cnt + 1
            Reference = Null2String(rsJournal_HDDet!jtype) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
            If Null2String(rsJournal_HDDet!jtype) = "VPJ" Then
                OUTBALANCE = OUTBALANCE + N2Str2Zero(rsJournal_HDDet!amounttopay)
            Else
                If Null2String(rsJournal_HDDet!jtype) = "VCJ" Or Null2String(rsJournal_HDDet!jtype) = "VDJ" Then
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!CM) - N2Str2Zero(rsJournal_HDDet!DM)))
                Else
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!CREDIT) - N2Str2Zero(rsJournal_HDDet!DEBIT)))
                End If
            End If

            xlSheet1.Cells(6 + cnt, "A") = Null2String(rsJournal_HDDet!jdate)
            xlSheet1.Cells(6 + cnt, "B") = Reference
            xlSheet1.Cells(6 + cnt, "C") = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
            xlSheet1.Cells(6 + cnt, "D") = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
            xlSheet1.Cells(6 + cnt, "E") = ToDoubleNumber(OUTBALANCE)

            TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
            TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
            rsJournal_HDDet.MoveNext
        Loop

    Else
        ' Do Nothing
    End If

    ' Set NARD = xlSheet1.Range(10 + cnt + 1, "E")

    '    NARD.Font.Size = 11
    '    NARD.Font.Bold = True
    '    NARD.Font.Underline = True

    xlSheet1.Cells(6 + cnt + 1, "E") = ToDoubleNumber(TUTAL_BALANCE + OUTBALANCE)
    xlSheet1.Cells(6 + cnt + 1, "D") = ToDoubleNumber(TUTAL_CREDIT)
    xlSheet1.Cells(6 + cnt + 1, "C") = ToDoubleNumber(TUTAL_DEBIT)
    xlSheet1.Cells(6 + cnt + 1, "B") = "TOTAL"

    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error GoTo Errorcode:

    Frame2.ZOrder 0
    On Error Resume Next
    TextSearch.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsVENDOR.MoveNext
    If rsVENDOR.EOF Then
        rsVENDOR.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsVENDOR.MovePrevious
    If rsVENDOR.BOF Then
        rsVENDOR.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    Dim Ans                                                           As String
    On Error GoTo Errorcode:
    '    If Function_Access(LOGID, "Acess_Print", "VENDOR LEDGER") = False Then Exit Sub
    '    Screen.MousePointer = 11
    '    rptVendor.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    '    rptVendor.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '    PrintReport rptVendor, AMIS_REPORT_PATH & "Vendorledger.rpt", "{all_vendor_table.code} = '" & txtCode.Text & "'", 1
    '    Screen.MousePointer = 0
    Ans = MsgBox("Are you sure do you want to print this Ledger", vbQuestion + vbYesNo)
    If Ans = vbYes Then
        PrintToExcel
        LogAudit "V", "VENDORS LEDGER", txtcode
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub
Private Sub FillGrid()
    Dim rsVendors                                                     As ADODB.Recordset
    lstVendor.Enabled = False
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendors = New ADODB.Recordset
    'Set rsVendors = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD INNER JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ' ORDER BY dbo.ALL_Vendor.NameofVendor")
    Set rsVendors = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Vendor.NameofVendor,dbo.ALL_Vendor.Code, dbo.ALL_Vendor.ID FROM dbo.AMIS_Journal_HD left outer JOIN dbo.ALL_Vendor ON dbo.AMIS_Journal_HD.VendorCode = dbo.ALL_Vendor.Code WHERE dbo.AMIS_Journal_HD.JType = 'APJ' OR dbo.AMIS_Journal_HD.JType = 'CDJ' OR dbo.AMIS_Journal_HD.JType = 'VPJ' ORDER BY dbo.ALL_Vendor.NameofVendor")
    If Not (rsVendors.EOF And rsVendors.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendors
        lstVendor.Refresh
        lstVendor.Enabled = True
        lstVendor.Enabled = True
    Else
        lstVendor.Enabled = False
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    TextSearch.Text = "": Frame2.ZOrder 1
    InitMemVars
    FillGrid
    StoreMemvars
    FillGrids
    Screen.MousePointer = 0
End Sub

Private Sub grdAccountsLedger_DblClick()
    grdAccountsLedger.Row = grdAccountsLedger.Row
    grdAccountsLedger.Col = 1
    Dim VARVOUCHERNO                                                  As String
    If Left(grdAccountsLedger.Text, 3) = "APJ" Then
        JOURNALTYPE = "APJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CDJ" Then
        JOURNALTYPE = "CDJ"
    ElseIf Left(grdAccountsLedger.Text, 2) = "SJ" Then
        JOURNALTYPE = "SJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CRJ" Then
        JOURNALTYPE = "CRJ"
    ElseIf Left(grdAccountsLedger.Text, 2) = "GJ" Then
        JOURNALTYPE = "GJ"
    Else
        JOURNALTYPE = Left(grdAccountsLedger.Text, 3)
    End If
    VARVOUCHERNO = Right(grdAccountsLedger.Text, 6)
    Screen.MousePointer = 11
    If JOURNALTYPE = "VPJ" Then
        On Error Resume Next
        Unload frmAMISVendorAPOpening
        frmAMISVendorAPOpening.Show
        frmAMISVendorAPOpening.StoreSearch (VARVOUCHERNO)
    ElseIf JOURNALTYPE = "VDJ" Then
        'MsgBox "Please open Vendor adjustment Instead", vbInformation, "Info"
        JOURNALTYPE = "VDJ"
        On Error Resume Next
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
    ElseIf JOURNALTYPE = "VCJ" Then
        MsgBox "Please open Vendor adjustment Instead", vbInformation, "Info"
    Else
        On Error Resume Next
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub lstVendor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstVendor
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstVendor_GotFocus()
    rsVENDOR.Bookmark = rsFind(rsVENDOR.Clone, "NameOfVendor", lstVendor.SelectedItem).Bookmark
    StoreMemvars
End Sub

Private Sub lstVendor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsVENDOR.Bookmark = rsFind(rsVENDOR.Clone, "NameOfVendor", lstVendor.SelectedItem).Bookmark
    StoreMemvars
End Sub

Private Sub lstVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        TextSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    'On Error Resume Next
    If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Frame2.ZOrder 0
    If KeyCode = vbKeyDown Then
        If lstVendor.ListItems.Count > 0 And lstVendor.Enabled = True Then
            lstVendor.SetFocus
        End If
    End If
End Sub

