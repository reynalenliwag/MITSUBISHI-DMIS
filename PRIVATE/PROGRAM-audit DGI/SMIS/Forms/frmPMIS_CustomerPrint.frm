VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSMIS_CustomerPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master File"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "frmPMIS_CustomerPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   1
      Left            =   2340
      MouseIcon       =   "frmPMIS_CustomerPrint.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_CustomerPrint.frx":101C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1320
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1170
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   1620
      MouseIcon       =   "frmPMIS_CustomerPrint.frx":1467
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_CustomerPrint.frx":15B9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1320
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   345
      Index           =   1
      Left            =   585
      TabIndex        =   0
      Top             =   855
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20578305
      CurrentDate     =   40693
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   345
      Index           =   0
      Left            =   585
      TabIndex        =   1
      Top             =   495
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20578305
      CurrentDate     =   40664
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   270
      TabIndex        =   5
      Top             =   945
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Left            =   -90
      TabIndex        =   4
      Top             =   630
      Width           =   765
   End
End
Attribute VB_Name = "frmSMIS_CustomerPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HD                                           As ADODB.Recordset

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0:
        printCustomer
    Case 1:
        Unload Me
End Select
End Sub

Private Sub opt_Click(Index As Integer)
Select Case Index
    Case 0:
        dtp(1).Enabled = True
        dtp(0).Enabled = True
        
    Case 1:
        dtp(1).Enabled = False
        dtp(0).Enabled = False
       
    End Select
End Sub
Private Sub printCustomer()
        If (dtp(0).Value > dtp(1).Value) Or IsDate(dtp(0).Value) = False Or IsDate(dtp(1).Value) = False Then
            MsgSpeechBox "Error In From and To date"
            Exit Sub
        End If

        Dim FDate                                      As Date
        Dim TDate                                      As Date
        dtp(0).Value = Format(dtp(0).Value, "Short Date")
        dtp(1).Value = Format(dtp(1).Value, "Short Date")

        FDate = CDate(dtp(0).Value)
        TDate = CDate(dtp(1).Value)

        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select ENTRY_DATE from ALL_Customer_Table where (ENTRY_DATE >= '" & FDate & "' AND ENTRY_DATE <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Screen.MousePointer = 11

        If Not RSORD_HD.EOF And Not RSORD_HD.EOF Then
            CrystalReport1.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            CrystalReport1.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            CrystalReport1.Formulas(12) = "mindate = '" & FDate & "'"
            CrystalReport1.Formulas(11) = "maxdate = '" & TDate & "'"
            PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "CustomerIsRange.rpt", "{ALL_Customer_Table.ENTRY_DATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ALL_Customer_Table.ENTRY_DATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
        Else
            ShowNoRecord
        End If

        Screen.MousePointer = 0
        
ErrorCode:
        ShowVBError
    

End Sub

Private Sub Form_Load()
dtp(1).Value = Now
dtp(0).Value = Now
End Sub

