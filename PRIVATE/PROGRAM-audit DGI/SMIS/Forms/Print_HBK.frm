VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmSMIS_Report_Print_HBK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Process..."
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Print_HBK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4425
   Begin VB.PictureBox picDebitMemo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3810
      TabIndex        =   8
      Top             =   990
      Visible         =   0   'False
      Width           =   3840
      Begin VB.TextBox txtDebitMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         MaxLength       =   6
         TabIndex        =   10
         ToolTipText     =   "Input Debit Memo Serial Number"
         Top             =   750
         Width           =   2445
      End
      Begin wizButton.cmd cmdPrintDebitmemo 
         Height          =   435
         Left            =   960
         TabIndex        =   11
         Top             =   1380
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0E42
      End
      Begin wizButton.cmd cmd7 
         Height          =   435
         Left            =   2070
         TabIndex        =   12
         Top             =   1380
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0E5E
      End
      Begin VB.Label Label3 
         Caption         =   "DEBIT MEMO #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   630
         TabIndex        =   9
         Top             =   450
         Width           =   1995
      End
   End
   Begin VB.PictureBox picGatePass 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   6420
      ScaleHeight     =   1065
      ScaleWidth      =   4110
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   4110
      Begin wizButton.cmd cmd2 
         Height          =   435
         Left            =   1620
         TabIndex        =   14
         Top             =   1680
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0E7A
      End
      Begin wizButton.cmd cmd3 
         Height          =   435
         Left            =   2730
         TabIndex        =   15
         Top             =   1680
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0E96
      End
      Begin wizButton.cmd cmdClearance 
         Height          =   435
         Left            =   300
         TabIndex        =   24
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   1650
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Clearance Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0EB2
      End
      Begin wizButton.cmd cmdDR 
         Height          =   435
         Left            =   240
         TabIndex        =   25
         Top             =   540
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Dealers Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0ECE
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   -30
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   -60
      Width           =   6360
      Begin Crystal.CrystalReport rptPrint 
         Left            =   390
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin wizButton.cmd cmdCreditMemo 
         Height          =   435
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Release Order"
         Top             =   1650
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Credit Memo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0EEA
      End
      Begin wizButton.cmd cmdDebitMemo 
         Height          =   435
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2135
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Debit Memo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F06
      End
      Begin wizButton.cmd cmdVDR 
         Height          =   435
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Vehicle Delivery Report"
         Top             =   680
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Vehicle Delivery Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F22
      End
      Begin wizButton.cmd cmdVI 
         Height          =   435
         Left            =   150
         TabIndex        =   1
         ToolTipText     =   "Vehicle Invoice"
         Top             =   195
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Vehicle &Invoice"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F3E
      End
      Begin wizButton.cmd cmdExit 
         Height          =   435
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   4560
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F5A
      End
      Begin wizButton.cmd cmd1 
         Height          =   435
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Gate Pass"
         Top             =   1165
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Gate Pass"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F76
      End
      Begin wizButton.cmd cmdReleaseOrder 
         Height          =   435
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2620
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Release Order"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0F92
      End
      Begin wizButton.cmd cmd4 
         Height          =   435
         Left            =   150
         TabIndex        =   21
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   3105
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Warranty"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0FAE
      End
      Begin wizButton.cmd cmd6 
         Height          =   435
         Left            =   150
         TabIndex        =   22
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   4075
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Signatories"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0FCA
      End
      Begin wizButton.cmd cmd8 
         Height          =   435
         Left            =   150
         TabIndex        =   23
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   3590
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Extended Warranty"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":0FE6
      End
   End
   Begin VB.PictureBox picCreditMemo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3810
      TabIndex        =   16
      Top             =   1170
      Visible         =   0   'False
      Width           =   3840
      Begin VB.TextBox txtCreditMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   630
         MaxLength       =   6
         TabIndex        =   18
         ToolTipText     =   "Input Credit Memo Serial Number"
         Top             =   630
         Width           =   2415
      End
      Begin wizButton.cmd cmdPrintCreditMemo 
         Height          =   435
         Left            =   960
         TabIndex        =   19
         Top             =   1260
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":1002
      End
      Begin wizButton.cmd cmd5 
         Height          =   435
         Left            =   2070
         TabIndex        =   20
         Top             =   1260
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print_HBK.frx":101E
      End
      Begin VB.Label Label4 
         Caption         =   "CREDIT MEMO #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   630
         TabIndex        =   17
         Top             =   330
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_Print_HBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                          As ADODB.Recordset
Dim rsPurchAgree                        As ADODB.Recordset
Public GM                               As String
Public IGNKEYNO                         As String
Public vi_no As String
Private Sub cmd1_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    LoadSignatories ("GATE PASS")
    'rptPrint.Formulas(0) = "GuardOnDuty=" & N2Str2Null(txtGatePassGuardOnDuty)
    'rptPrint.Formulas(1) = "TimeOut=" & N2Str2Null(txtGatePassTimeOut)
    'rptPrint.Formulas(0) = "FinancingManager=" & N2Str2Null(FinancingManager)
    rptPrint.Formulas(0) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(1) = "ReleasedBy='" & Null2String(SalesDispatcher) & "'"
    
    If COMPANY_CODE = "HBK" Then
        rptPrint.Formulas(2) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
        rptPrint.Formulas(3) = "D_RELBY='" & Null2String(SalesDispatcherDesig) & "'"
    End If
    
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "GatePass.rpt", "{SMIS_SalesOrder.Prodno}='" & PRODUCTNO & "' and {SMIS_SalesOrder.code}='" & CUSCODE & "'", DMIS_REPORT_Connection, 1
    cmd3_Click
    Screen.MousePointer = 0
End Sub

Private Sub cmd3_Click()
    picMain.Visible = True
    picGatePass.Visible = False
End Sub

Private Sub cmd4_Click()
    PRINTWARRANTYEXCEL IGNKEYNO
End Sub
Sub PRINTEXWARRANTYEXCEL(IGNKEYNO)
    On Error GoTo ErrorCode
    Dim xlApp                           As Excel.Application
    Dim xlBook                          As Excel.Workbook
    Dim xlSheet                         As Excel.Worksheet
    Set xlApp = New Excel.Application
    If Len(Dir(App.Path & "\SMIS_EXCEL\EXTENDED WARRANTY.xls")) = 0 Then
    MsgBox "Missing Template File .", vbInformation
    
    Exit Sub
    End If
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\SMIS_EXCEL\EXTENDED WARRANTY.xls")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                         As ADODB.Recordset
    Dim vmodel                          As String
    Dim i                               As Integer
    Dim j                               As Integer
    Dim rsMrr                           As ADODB.Recordset
    Dim rsCountProspect                 As ADODB.Recordset

    Set rsModel = gconDMIS.Execute("Select * from SMIS_SALESORDER where IGNKEY_NO='" & IGNKEYNO & "'")

    Set rsMrr = gconDMIS.Execute("Select * from SMIS_MRRINV where ignkey='" & IGNKEYNO & "'")
    If Not rsModel.EOF Or Not rsModel.BOF Then

        xlSheet.Cells(2, "G") = Space(12) & Null2String(rsModel("CUSTNAME"))
        xlSheet.Cells(4, "I") = Null2String(rsModel("HOMEADDRESS"))

        Dim myChar
        Dim vinchar

        If Null2String(rsModel("VINO")) <> "" Then
            myChar = Null2String(rsModel("VINO"))
            For i = 1 To Len(myChar)
                vinchar = vinchar & Mid(myChar, i, 1) & Space(5)
            Next

        End If
        'Stop
        xlSheet.Cells(8, "A") = Space(8) & vinchar
        xlSheet.Cells(10, "B") = Null2String(rsModel("MODEL"))
        If Not rsMrr.EOF Or Not rsMrr.BOF Then
            xlSheet.Cells(10, "D") = Null2String(rsMrr("yeer"))
            xlSheet.Cells(10, "E") = Null2String(rsMrr("ENGINENO"))
        End If
        xlSheet.Cells(16, "C") = Null2String(rsModel("SALESPRICE"))
        xlSheet.Cells(26, "F") = Null2String(rsModel("HOMEADDRESS"))
        If IsDate(rsModel("DateReleased")) = True Then
            xlSheet.Cells(27, "D") = Space(15) & Day(rsModel("DateReleased")) & Space(9) & Month(rsModel("DateReleased")) & Space(6) & Year(rsModel("DateReleased"))
            xlSheet.Cells(27, "I") = Space(15) & Day(rsModel("DateReleased")) & Space(9) & Month(rsModel("DateReleased")) & Space(6) & Year(rsModel("DateReleased"))
        End If
        xlApp.Visible = True
        Set xlApp = Nothing
    End If
    Exit Sub
ErrorCode:
    MsgBox Err.Description
    Err.Clear
End Sub

Sub PRINTWARRANTYEXCEL(xYear)
    On Error GoTo ErrorCode
    Dim xlApp                           As Excel.Application
    Dim xlBook                          As Excel.Workbook
    Dim xlSheet                         As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\SMIS_EXCEL\warranty.xls")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                         As ADODB.Recordset
    Dim vmodel                          As String
    Dim i                               As Integer
    Dim j                               As Integer
    Dim rsCountProspect                 As ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select * from SMIS_SALESORDER where IGNKEY_NO='" & IGNKEYNO & "'")
    If Not rsModel.EOF Or Not rsModel.BOF Then
        xlSheet.Cells(1, 1) = Null2String(rsModel("CUSTNAME"))
        xlSheet.Cells(3, 1) = Null2String(rsModel("HOMEADDRESS"))

        xlSheet.Cells(9, 1) = Null2String(rsModel("INVOICEDDATE"))
        xlSheet.Cells(9, 2) = Null2String(rsModel("MODELDESCRIPTION"))
        xlSheet.Cells(9, "D") = Null2String(rsModel("IGNKEY_NO"))
        xlSheet.Cells(11, "A") = Null2String(rsModel("VINO"))
        xlSheet.Cells(13, "A") = Null2String(rsModel("ENGINENO"))
        xlSheet.Cells(13, "D") = Null2String(rsModel("COLOR"))



        xlSheet.Cells(22, "B") = Null2String(rsModel("CUSTNAME"))
        xlSheet.Cells(22, 4) = Null2String(rsModel("DATERELEASED"))
        xlApp.Visible = True
        Set xlApp = Nothing
    End If
    Exit Sub
ErrorCode:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmd6_Click()
    frmSMIS_Files_Signatories.Show 1
End Sub

Private Sub cmd8_Click()
    PRINTEXWARRANTYEXCEL IGNKEYNO
End Sub

Private Sub cmdClearance_Click()
    rptPrint.Reset
    Screen.MousePointer = 11
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "clearance.rpt", "{customer.cuscde} = '" & CUSCODE & "' AND {PurchAgree.ProdNo} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmd5_Click()
    ShowHidePictureBox2 picCreditMemo, False, picMain
End Sub

Private Sub cmd7_Click()
    ShowHidePictureBox2 picDebitMemo, False, picMain
End Sub

Private Sub cmdCreditMemo_Click()
    Dim TEMPRS                          As ADODB.Recordset
    rptPrint.Reset
    Set TEMPRS = gconDMIS.Execute("select Count(*) from SMIS_MrrInv_Detail where  IgnKeyNo='" & IGNKEYNO & "'")
    If TEMPRS(0).Value = 0 And rsPurchAgree!DISCOUNT = 0 Then
        MsgBox " There are No Record For This Transaction", vbInformation
        Exit Sub
    End If
    Dim RSCREDIT                        As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("select creditmemo from smis_salesorder where ignkey_no='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("CREDITMEMO").Value) = True Then
        txtCreditMemo = (GenerateCode("SMIS_SALESORDER", "CREDITMEMO", "000000"))
    Else
        txtCreditMemo = RSCREDIT("CREDITMEMO").Value
    End If
    ShowHidePictureBox2 picCreditMemo, True, picMain


End Sub
Private Sub cmdDebitMemo_Click()
    rptPrint.Reset
    Dim RSCREDIT                        As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("SELECT DEBITMEMO FROM SMIS_SALESORDER WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("DEBITMEMO").Value) = True Then
        txtDebitMemo = (GenerateCode("SMIS_SALESORDER", "DEBITMEMO", "000000"))
    Else
        txtDebitMemo = RSCREDIT("DEBITMEMO").Value
    End If
    Set RSCREDIT = Nothing
    ShowHidePictureBox2 picDebitMemo, True, picMain
End Sub

Private Sub cmdDR_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "dealers.rpt", "{customer.code} = '" & CUSCODE & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrintCreditMemo_Click()
    Screen.MousePointer = 11
    '''''''AXP063020071200
    Dim lng                             As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE CREDITMEMO=" & N2Str2Null(txtCreditMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!CREDITMEMO)) <> UCase(txtCreditMemo) Then
        MessagePop RecSaveWarning, "Duplicate Record", "Credit Memo Number Already Exist"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET CREDITMEMO=" & N2Str2Null(txtCreditMemo) & " WHERE VI_NO='" & vi_no & "'")
    rsRefresh
    rptPrint.Reset
    rptPrint.WindowShowPrintBtn = True
    
    'JUN 07/25/2008
        LoadSignatories ("CREDIT MEMO")
        rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
        rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
        rptPrint.Formulas(2) = "SalesApproved='" & Null2String(ApprovedBy) & "'"
        rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            
        If COMPANY_CODE = "HBK" Then
            rptPrint.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
            rptPrint.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
            rptPrint.Formulas(6) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
            rptPrint.Formulas(7) = "D_GM='" & Null2String(GeneralManagerDesig) & "'"
        End If
    
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "({purchagree.vi_no} = '" & vi_no & "')", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    ShowHidePictureBox2 picCreditMemo, False, picMain
End Sub

Private Sub cmdPrintDebitmemo_Click()
    Screen.MousePointer = 11
    '''''''AXP063020071200
    Dim lng                             As Integer
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEBITMEMO=" & N2Str2Null(txtDebitMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!DEBITMEMO)) <> UCase(txtDebitMemo) Then
        MessagePop RecSaveWarning, "DUPLICATE RECORD", "DEBIT MEMO NUMBER ALREADY EXIST"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET DEBITMEMO=" & N2Str2Null(txtDebitMemo) & " WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    rsRefresh
    LoadSignatories ("DEBIT MEMO")
    rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "receivedby='" & Null2String(FinancingManager) & "'"
    
    If COMPANY_CODE = "HBK" Then
        rptPrint.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
        rptPrint.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
        rptPrint.Formulas(6) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
    End If
    
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "DEBITMEMO.RPT", "{PurchAgree.IGNKEY_NO} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    ShowHidePictureBox2 picDebitMemo, False, picMain
End Sub

Private Sub cmdReleaseOrder_Click()
    rptPrint.Reset
    'If Not rsCustomer.EOF And Not rsCustomer.BOF Then
    '    If Null2String(rsPurchAgree!Term) = "COD" Then
    
    'UPDATED BY: JUN 07/25/2005
    
    If COMPANY_CODE = "HBK" Then
    
    Dim rsDscount As ADODB.Recordset
    Dim sumDis As Double

    sumDis = 0
    Set rsDscount = gconDMIS.Execute("Select COST from SMIS_MrrInv_Detail where IGNKEYNO = '" & frmSMIS_Trans_VehicleInvoice.txtVehicleConductionSticker & "' and Description <> 'DISCOUNT' ")
    If Not rsDscount.BOF And Not rsDscount.EOF Then
        Do While Not rsDscount.EOF
            sumDis = ToDoubleNumber(rsDscount!Cost) + sumDis
            rsDscount.MoveNext
        Loop
    End If
    
    End If
    
    Screen.MousePointer = 11
    rptPrint.Reset
    rptPrint.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrint.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
    If COMPANY_CODE = "HBK" Then
        If ToDoubleNumber(sumDis) = 0 Then
            sumDis = 0
            rptPrint.Formulas(44) = "TotalDiscount = '" & sumDis & "'"
        Else
            rptPrint.Formulas(51) = "TotalDiscount = '" & sumDis & "'"
        End If
    End If
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "ReleaseOrder.rpt", "{purchagree.ignKey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    '     Else
    '         Screen.MousePointer = 11
    '         PrintSQLReport rptPrint, SMIS_REPORT_PATH & "releaseFI.rpt", "{customer.code} = '" & CusCode & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    '         Screen.MousePointer = 0
    '     End If
    'End If
End Sub
Private Sub cmdVDR_Click()
    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
        rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
        rptPrint.Formulas(2) = "SalesApproved='" & Null2String(ApprovedBy) & "'"
        rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
        
        If COMPANY_CODE = "HBK" Then
            rptPrint.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
            rptPrint.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
            rptPrint.Formulas(6) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
        End If
                
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vdr.rpt", "{customer.cuscde} = '" & CUSCODE & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

'Upating Code       : AXP-0707200712:44
'Upating Code       : AXP-0707200712:44
Private Sub cmdVI_Click()
    On Error GoTo ErrorCode:
    rptPrint.Reset
    LoadSignatories ("SALES INVOICE")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then

        If Null2String(rsPurchAgree!TERM) = "COD" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM='" & Null2String(GeneralManager) & "'"
            rptPrint.Formulas(1) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(2) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(3) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
            
            If COMPANY_CODE = "HBK" Then
                rptPrint.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
                rptPrint.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
                rptPrint.Formulas(6) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
            End If
            
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi.rpt", "{customer.CUSCDE} = '" & CUSCODE & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM='" & Null2String(GeneralManager) & "'"
            rptPrint.Formulas(1) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(2) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(3) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
            
            If COMPANY_CODE = "HBK" Then
                rptPrint.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
                rptPrint.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
                rptPrint.Formulas(6) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
            End If
            
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi_fin.rpt", "{customer.CUSCDE} = '" & CUSCODE & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    rsRefresh
    'Dim RSSIGNATORIES
    '  Set RSSIGNATORIES = New ADODB.Recordset
    '        RSSIGNATORIES.Open "select * from SMIS_Signatories", gconDMIS, adOpenForwardOnly, adLockReadOnly
    '        If Not RSSIGNATORIES.EOF And Not RSSIGNATORIES.BOF Then
    '            GM = Null2String(RSSIGNATORIES!GeneralManager)
    '        End If
    '    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
    '
    '
    '        If Null2String(rsPurchAgree!DateReleased) = "" Then
    '            cmdCreditMemo.Enabled = False
    '            cmdVDR.Enabled = False
    '        Else
    '            cmdCreditMemo.Enabled = True
    '            cmdVDR.Enabled = True
    '        End If
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from ALL_CUSTMASTER_SMIS where code = '" & CUSCODE & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If rsCustomer.BOF And rsCustomer.EOF Then
        MsgSpeechBox "Error Encountered! Empty Customer Record!"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GM = ""
    IGNKEYNO = ""
End Sub

Sub rsRefresh()
    Set rsPurchAgree = New ADODB.Recordset

    rsPurchAgree.Open "select * from SMIS_SALESORDER where code = '" & CUSCODE & "' AND IGNKEY_NO ='" & IGNKEYNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub txtCreditMemo_LostFocus()
    txtCreditMemo = Format(txtCreditMemo, "000000")
End Sub

