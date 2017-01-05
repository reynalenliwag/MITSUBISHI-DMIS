VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_VehiclesInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Inventory Report"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ForeColor       =   &H8000000F&
   Icon            =   "Report_VehiclesInventory.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4965
   Begin VB.OptionButton Opt 
      Caption         =   "Select Date Ranged"
      Height          =   375
      Index           =   1
      Left            =   210
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Select Monthly"
      Height          =   375
      Index           =   2
      Left            =   210
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.PictureBox picDateRange 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   30
      ScaleHeight     =   1005
      ScaleWidth      =   4725
      TabIndex        =   13
      Top             =   810
      Width           =   4725
      Begin MSComCtl2.DTPicker datepFrom 
         Height          =   405
         Left            =   1470
         TabIndex        =   14
         Top             =   90
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53608449
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker datepTo 
         Height          =   405
         Left            =   1470
         TabIndex        =   15
         Top             =   540
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53608449
         CurrentDate     =   38216
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   690
         TabIndex        =   17
         Top             =   90
         Width           =   675
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   630
         Width           =   435
      End
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Select as of Date"
      Height          =   375
      Index           =   0
      Left            =   210
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Report_VehiclesInventory.frx":0E42
      Left            =   750
      List            =   "Report_VehiclesInventory.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   450
      Width           =   3555
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2430
      MouseIcon       =   "Report_VehiclesInventory.frx":0E46
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehiclesInventory.frx":0F98
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1920
      Width           =   885
   End
   Begin Crystal.CrystalReport rptInvent 
      Left            =   4320
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "DMC Purchases Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1560
      MouseIcon       =   "Report_VehiclesInventory.frx":13E3
      MousePointer    =   99  'Custom
      Picture         =   "Report_VehiclesInventory.frx":1535
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1920
      Width           =   885
   End
   Begin VB.PictureBox picAsofDate 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   1590
      ScaleHeight     =   885
      ScaleWidth      =   1575
      TabIndex        =   9
      Top             =   900
      Width           =   1575
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   -60
         Width           =   2505
      End
   End
   Begin VB.PictureBox picMonthly 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   720
      ScaleHeight     =   1215
      ScaleWidth      =   3855
      TabIndex        =   4
      Top             =   900
      Width           =   3855
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   60
         Width           =   2535
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   510
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   30
         TabIndex        =   8
         Top             =   570
         Width           =   510
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Select Your Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   90
      Width           =   3255
   End
End
Attribute VB_Name = "frmSMIS_Report_VehiclesInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset
Dim DEALER_TYPE

Sub VEHICLEINVENTORY_ASOFDATE()
    Dim FILTER
    
    Dim FDate As String
    
    FDate = CDate(txtDate.Text)
    If picAsofDate.Visible = True Then
        If IsDate(txtDate.Text) = True Then
            Screen.MousePointer = 11
            
            'JBF 1/28/2011: HCA request
            'gconDMIS.Execute "update SMIS_MrrInv set lastinvdate = '" & txtDate.Text & "' "
            'rptInvent.Formulas(0) = "mindate = '" & fdate & "'"
            'PrintSQLReport rptInvent, SMIS_REPORT_PATH & "unitinventory.rpt", "((({VEHICLE.PULLOUTDATE} <= Date(" & Year(txtDate) & "," & Month(txtDate) & "," & Day(txtDate) & ")) AND ({VEHICLE.PULLOUTDATE}  < DATE(" & Year(txtDate) & "," & Month(txtDate) & "," & Day(txtDate) & "))) AND ({VEHICLE.Status}) = 'P') ", DMIS_REPORT_Connection, 1
        
            Call inv_as_date
            
            rptInvent.PageZoom 90
            Screen.MousePointer = 0
            
            Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
        End If
    ElseIf picDateRange.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "unitinventory.rpt", "((({VEHICLE.PULLOUTDATE} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({VEHICLE.PULLOUTDATE} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & ")))) ", DMIS_REPORT_Connection, 1
        'COMMENTED BY JUN
        'REASON: FILTERING NOR FUNCTIONING WELL DUE TO DIFFERENCE OF {VEHICLE.PULLOUTDATE} AND {VEHICLE.DateReleased}
        'PrintSQLReport rptInvent, SMIS_REPORT_PATH & "unitinventory.rpt", "{VEHICLE.PULLOUTDATE} <= Date(" & Year(txtDate) & "," & Month(txtDate) & "," & Day(txtDate) & ") AND {VEHICLE.DateReleased} < DATE(" & Year(txtDate) & "," & Month(txtDate) & "," & Day(txtDate) & ")) AND ({VEHICLE.Status}) = 'P' and {VEHICLE.DateReleased} = false ", DMIS_REPORT_Connection, 1
        rptInvent.PageZoom 90
        Screen.MousePointer = 0
        
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 440
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    ElseIf picMonthly.Visible = True Then
     If cboMonth = "ALL" Then
            rptInvent.Formulas(3) = "ASOF=' For The Year of " & cboYear & "'"
            FILTER = " YEAR({VEHICLE.PULLOUTDATE})= " & cboYear.Text & " "
        Else
            rptInvent.Formulas(3) = "ASOF=' For The Month Of " & cboMonth & " " & cboYear & "'"
            FILTER = " YEAR({VEHICLE.PULLOUTDATE})=" & cboYear.Text & " AND MONTH({VEHICLE.PULLOUTDATE})=" & What_month(cboMonth.Text) & "  "
        End If

        PrintSQLReport rptInvent, SMIS_REPORT_PATH & "unitinventory.rpt", FILTER, DMIS_REPORT_Connection, 1
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 440
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
    Else
        MsgSpeechBox "Invalid Date!"
    End If
    Exit Sub

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Sub inv_as_date()

Dim RSTMP                                           As New ADODB.Recordset
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

        'cmd.NamedParameters = True
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "INVENTORY_AS_DATE"
        cmd.ActiveConnection = gconDMIS
        cmd.Parameters.Append cmd.CreateParameter("@TODATE", adDBDate, adParamInput, , txtDate.Text)
        Set RSTMP = cmd.Execute
        
        
        If Not (RSTMP.EOF And RSTMP.BOF) Then
                If Len(Dir(SMIS_REPORT_PATH & "Inventory.xlt")) = 0 Then
                    MessagePop InfoStop, "Error", "Inventory.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                Set xlApp = New Excel.Application
                Set xlBook = xlApp.Workbooks.Open(SMIS_REPORT_PATH & "Inventory.xlt")
                Set xlSheet = xlBook.Worksheets(1)
                
               xlSheet.Cells(1, "B") = COMPANY_NAME
               xlSheet.Cells(2, "B") = COMPANY_ADDRESS
               xlSheet.Cells(5, "B") = "As of " & txtDate.Text
               
               xlSheet.Range("A8").CopyFromRecordset RSTMP
               xlApp.Visible = True
                If Not xlBook Is Nothing Then
                    Set xlBook = Nothing
                    Set xlApp = Nothing
                End If
                Set xlApp = Nothing
            Else
                Call ShowNoRecord
            End If
            Set RSTMP = Nothing
            Screen.MousePointer = 0
End Sub




Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    Dim FILTER                                                        As String
    rptInvent.Reset
    rptInvent.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvent.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptInvent.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
    rptInvent.WindowShowSearchBtn = True

    If Combo1.Text = "VEHICLES INVENTORY AS OF DATE" Then
        
        VEHICLEINVENTORY_ASOFDATE
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY MODEL" Then
        If COMPANY_CODE = "HAS" Then
            'UPDATED BY: JUN
            'DATE UPDATED: 09272008 1:21
            'DESCRIPTION: FOR HAS REPORT ONLY THEY WANT A REPORT WHICH IS GROUP AND TOTAL BY VEHICLE VARIANT
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "vehstockHAS.rpt", "{VEHICLE.Status}='P' AND {VEHICLE.Released}=FALSE", DMIS_REPORT_Connection, 1
        Else
            If picAsofDate.Visible = True Then
                PrintSQLReport rptInvent, SMIS_REPORT_PATH & "vehstock.rpt", "{VEHICLE.Status}='P' AND {VEHICLE.Released}=FALSE", DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            ElseIf picDateRange.Visible = True Then
                PrintSQLReport rptInvent, SMIS_REPORT_PATH & "vehstock.rpt", "((({VEHICLE.PULLOUTDATE} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({VEHICLE.PULLOUTDATE} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) and {VEHICLE.Status}='P' AND {VEHICLE.Released}=FALSE) ", DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            End If
        End If
        
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY VEHICLE TYPE" Then
        Screen.MousePointer = 11
        rptInvent.WindowTitle = "VEHICLE MODEL LIST"
        
        If picAsofDate.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "VehiclesGroupList.rpt", "({SMIS_MRRINV.Status}='P'  AND {SMIS_MRRINV.Released}=FALSE)", DMIS_REPORT_Connection, 1
            
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            Screen.MousePointer = 0
        ElseIf picDateRange.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "VehiclesGroupList.rpt", "((({SMIS_MRRINV.PULLOUTDATE} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({SMIS_MRRINV.PULLOUTDATE} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) and {SMIS_MRRINV.Status}='P'  AND {SMIS_MRRINV.Released}=FALSE) ", DMIS_REPORT_Connection, 1
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
                Screen.MousePointer = 0
        End If
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY AGING" Then
        rptInvent.Formulas(3) = "ASOF=' As of " & txtDate & "'"
        
        If picAsofDate.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehilcesOpen.rpt", "({SMIS_MRRINV.Status}='P'  AND {SMIS_MRRINV.Released}=FALSE)", DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        ElseIf picDateRange.Visible = True Then
             PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehilcesOpen.rpt", "((({SMIS_MRRINV.PULLOUTDATE} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({SMIS_MRRINV.PULLOUTDATE} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) and {SMIS_MRRINV.Status}='P'  AND {SMIS_MRRINV.Released}=FALSE) ", DMIS_REPORT_Connection, 1
        
        End If
        
    ElseIf Combo1.Text = "VEHICLES ALLOCATION AS PER SALES ORDER" Then
        
        If picDateRange.Visible = True Then
           PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesAllocated.rpt", "((({MRR.PULLOUTDATE} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({MRR.PULLOUTDATE} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) and {SO.SOSTATUS} ='P' and  {MRR.ISTATUS}='A' )", DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
               
        ElseIf picMonthly.Visible = True Then
            If cboMonth = "ALL" Then
                rptInvent.Formulas(3) = "ASOF=' For The Year of " & cboYear & "'"
                FILTER = " YEAR({SO.DEYT})=" & cboYear & " AND {SO.SOSTATUS} ='P' and  {MRR.ISTATUS}='A'"
            Else
                rptInvent.Formulas(3) = "ASOF=' For The Month Of " & cboMonth & " " & cboYear & "'"
                FILTER = " YEAR({SO.DEYT})=" & cboYear & " AND MONTH({SO.DEYT})=" & What_month(cboMonth) & "  AND {SO.SOSTATUS} ='P'   AND  {MRR.ISTATUS}='A'"
            End If
        
        PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesAllocated.rpt", FILTER, DMIS_REPORT_Connection, 1
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 440
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        End If
    ElseIf Combo1.Text = "INVOICED VEHICLES REPORT" Then
        If picDateRange.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesAllocated.rpt", "((({SO.InvoicedDate} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({SO.InvoicedDate} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))))", DMIS_REPORT_Connection, 1
        
        ElseIf picMonthly.Visible = True Then
        
            If cboMonth.Text = "ALL" Then
                rptInvent.Formulas(3) = "ASOF=' For The Year " & cboYear & "'"
                FILTER = " YEAR({SO.InvoicedDate})=" & cboYear
            Else
                rptInvent.Formulas(3) = "ASOF=' For The Month of " & cboMonth & " " & cboYear & "'"
                FILTER = " YEAR({SO.InvoicedDate})=" & cboYear & " AND MONTH({SO.InvoicedDate})=" & What_month(cboMonth)
            End If
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesInvoiced.rpt", FILTER, DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    ElseIf Combo1.Text = "RELEASED VEHICLES REPORT" Then
        If picDateRange.Visible = True Then
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesReleased.rpt", "((({SO.DateReleased} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({SO.DateReleased} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))))", DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        ElseIf picMonthly.Visible = True Then
            If cboMonth.Text = "ALL" Then
                rptInvent.Formulas(3) = "ASOF=' FOR THE YEAR OF " & cboYear & "'"
                FILTER = " YEAR({SO.DateReleased})=" & cboYear
            Else
                rptInvent.Formulas(3) = "ASOF=' For The Month Of " & cboMonth & " " & cboYear & "'"
                FILTER = " YEAR({SO.DateReleased})=" & cboYear & " AND MONTH({SO.DateReleased})=" & What_month(cboMonth)
            End If
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesReleased.rpt", FILTER, DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
         End If
    ElseIf Combo1.Text = "TRANSFERRED UNIT REPORT" Then
            
        If picDateRange.Visible = True Then
        PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesTransferred.rpt", "((({ST.Deyt} <= Date(" & Year(datepFrom) & "," & Month(datepFrom) & "," & Day(datepFrom) & ")) AND ({ST.Deyt} <= Date(" & Year(datepTo) & "," & Month(datepTo) & "," & Day(datepTo) & "))) and {MRR.Released}=true and {MRR.ISTATUS}='T' )", DMIS_REPORT_Connection, 1
       
        ElseIf picMonthly.Visible = True Then
            If cboMonth.Text = "ALL" Then
                rptInvent.Formulas(3) = "ASOF=' For The Year of " & cboYear & "'"
                FILTER = "  {MRR.Released}=true and {MRR.ISTATUS}='T' AND YEAR({ST.Deyt})=" & cboYear
            Else
                rptInvent.Formulas(3) = "ASOF=' For The Month Of " & cboMonth & " " & cboYear & "'"
                FILTER = "  {MRR.Released}=true and {MRR.ISTATUS}='T' AND YEAR({ST.Deyt})=" & cboYear & " AND MONTH({ST.Deyt})=" & What_month(cboMonth)
            End If
    
    
            PrintSQLReport rptInvent, SMIS_REPORT_PATH & "INVENTORY\VehiclesTransferred.rpt", FILTER, DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 440
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "VEHICLE INVENTORY REPORT", "", "", "", Combo1 & " " & txtDate, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    Else
        MsgBox "Please Select Report Type from the list", vbInformation
        On Error Resume Next
        Combo1.SetFocus
    End If


    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Combo1_Change()

    If Combo1.Text = "VEHICLES INVENTORY AS OF DATE" Then
        picMonthly.Visible = False: picAsofDate.Visible = True
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY MODEL" Then
        picMonthly.Visible = False: picAsofDate.Visible = True
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY VEHICLE TYPE" Then
        picMonthly.Visible = False: picAsofDate.Visible = True
    ElseIf Combo1.Text = "VEHICLES ON STOCK - BY AGING" Then
        picMonthly.Visible = False: picAsofDate.Visible = True
    Else
        picMonthly.Visible = True: picAsofDate.Visible = False

    End If
End Sub

Private Sub Combo1_Click()
    Combo1_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE INVENTORY REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE INVENTORY REPORT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtDate.Text = LOGDATE
    SetComboWidth Combo1, 300
    With Combo1
        .AddItem "VEHICLES INVENTORY AS OF DATE"
        .AddItem "VEHICLES ON STOCK - BY MODEL"
        .AddItem "VEHICLES ON STOCK - BY VEHICLE TYPE"
        .AddItem "VEHICLES ON STOCK - BY AGING"
        .AddItem "VEHICLES ALLOCATION AS PER SALES ORDER"
        .AddItem "INVOICED VEHICLES REPORT"
        .AddItem "RELEASED VEHICLES REPORT"
        .AddItem "TRANSFERRED UNIT REPORT"
        .ListIndex = 0
    End With
    fillcbomonth cboMonth
    cboMonth.AddItem "ALL"
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    picDateRange.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub Opt_Click(Index As Integer)
If Opt(0).Value = True Then
    picDateRange.Visible = False
    picAsofDate.Visible = True
    picMonthly.Visible = False
    Opt(0).BackColor = &HFFFF80
    Opt(1).BackColor = &H8000000F
ElseIf Opt(1).Value = True Then
    picDateRange.Visible = True
    picAsofDate.Visible = False
    picMonthly.Visible = False
    Opt(0).BackColor = &HFFFF80
    Opt(2).BackColor = &H8000000F
Else
    picMonthly.Visible = True
    picAsofDate.Visible = False
    picDateRange.Visible = False
    Opt(1).BackColor = &HFFFF80
    Opt(0).BackColor = &H8000000F
 
    
End If
End Sub

