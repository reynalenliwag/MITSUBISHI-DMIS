VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSLoadEstimateToRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Estimate to Repair Order"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmLoadEstimateToRO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10425
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   3180
      ScaleHeight     =   4095
      ScaleWidth      =   7155
      TabIndex        =   9
      Top             =   30
      Width           =   7185
      Begin Crystal.CrystalReport rptRepairOrder 
         Left            =   240
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtdate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox TXTWrite 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2040
         Width           =   3795
      End
      Begin VB.TextBox txtKM 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   6390
         MouseIcon       =   "FrmLoadEstimateToRO.frx":014A
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadEstimateToRO.frx":029C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtAcct_No 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   17
         Top             =   1230
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtEstimateno 
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
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1725
      End
      Begin VB.TextBox txtCustomer 
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
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   5385
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1170
         Width           =   2835
      End
      Begin VB.TextBox txtPlanteNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   1755
      End
      Begin VB.TextBox txtROno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   5130
         TabIndex        =   12
         Top             =   360
         Width           =   1875
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5670
         MouseIcon       =   "FrmLoadEstimateToRO.frx":05DA
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadEstimateToRO.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Edit Selected Record"
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton cmdProces 
         Caption         =   "&Upload"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4950
         MouseIcon       =   "FrmLoadEstimateToRO.frx":0A88
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadEstimateToRO.frx":0BDA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Process Upload"
         Top             =   3120
         Width           =   735
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
         Height          =   915
         Left            =   4230
         MouseIcon       =   "FrmLoadEstimateToRO.frx":0E75
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadEstimateToRO.frx":0FC7
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Print this Record"
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estimate By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   570
         TabIndex        =   30
         Top             =   2130
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Odometer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   690
         TabIndex        =   29
         Top             =   1710
         Width           =   840
      End
      Begin VB.Line Line1 
         X1              =   -30
         X2              =   7140
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estimate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   510
         TabIndex        =   25
         Top             =   450
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   24
         Top             =   930
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3570
         TabIndex        =   23
         Top             =   1290
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Plate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   825
         TabIndex        =   22
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "New R/O No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4020
         TabIndex        =   21
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Date Estimate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   405
         TabIndex        =   20
         Top             =   2550
         Visible         =   0   'False
         Width           =   1125
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   8265
         _Version        =   655364
         _ExtentX        =   14579
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ESTIMATE INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   30
      ScaleHeight     =   4095
      ScaleWidth      =   3105
      TabIndex        =   3
      Top             =   30
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "Estimate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   750
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   6
         Top             =   750
         Width           =   1065
      End
      Begin MSComctlLib.ListView lstEstimate 
         Height          =   3015
         Left            =   30
         TabIndex        =   7
         Top             =   1020
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmLoadEstimateToRO.frx":132D
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Estimate No"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtKeyword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   30
         TabIndex        =   5
         Top             =   330
         Width           =   3045
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4245
         _Version        =   655364
         _ExtentX        =   7488
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "SEARCH"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox Frame3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   3300
      ScaleHeight     =   1215
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   2100
      Width           =   4185
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   3780
         Top             =   -60
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4245
         _Version        =   655364
         _ExtentX        =   7488
         _ExtentY        =   450
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin VB.Label A 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Estimate Order Available"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   270
         TabIndex        =   1
         Top             =   570
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmCSMSLoadEstimateToRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUPLOAD                                           As ADODB.Recordset

Function FindSAName(VCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * from CSMS_VW_EMPNO WHERE CODE = '" & VCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindSAName = RSTMP!NAYM
    End If
    Set RSTMP = Nothing
End Function

Function GetNewROno(XXX As Variant)
    Dim rsNewRO                                        As ADODB.Recordset
    Set rsNewRO = New ADODB.Recordset
    Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc")
    If Not rsNewRO.EOF And Not rsNewRO.BOF Then
        GetNewROno = Format(NumericVal(Mid$(rsNewRO!rep_OR, 3, 8)) + 1, "R-00000000")
    Else
        GetNewROno = "R-00000001"
    End If
    Set rsNewRO = Nothing
End Function

Sub GenerateNewRONO()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim VRO                                            As Long
    Set RSTMP = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR ORDER BY REP_OR DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveFirst
        VRO = Mid(RSTMP!rep_OR, 3, 8) + 1
        txtROno.Text = "R-" & Format(VRO, "00000000")
    Else
        txtROno.Text = "R-00000001'"
    End If

    Set RSTMP = Nothing
End Sub

Sub FillGrid()
    Dim RSTMP                                          As New ADODB.Recordset
    lstEstimate.Enabled = False
    lstEstimate.Sorted = False: lstEstimate.ListItems.Clear

    Set RSTMP = gconDMIS.Execute("select ESTIMATENO,NIYM, ID from CSMS_REPOR WHERE TRANSTYPE = 'E' Order by ID")

    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Listview_Loadval Me.lstEstimate.ListItems, RSTMP
        lstEstimate.Refresh
        lstEstimate.Enabled = True
    End If

    Set RSTMP = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP                                          As New ADODB.Recordset
    
    lstEstimate.Sorted = False: lstEstimate.ListItems.Clear
    lstEstimate.Enabled = False
    XXX = Replace(LTrim(RTrim(XXX)), "'", "")

    If Option2.Value = True Then
        Set RSTMP = gconDMIS.Execute("select ESTIMATENO,NIYM, ID from CSMS_REPOR where TRANSTYPE = 'E' AND NIYM Like '%" & XXX & "%' ORDER BY ID")
    Else
        Set RSTMP = gconDMIS.Execute("select ESTIMATENO,NIYM, ID from CSMS_REPOR where TRANSTYPE = 'E' AND ESTIMATENO Like '%" & XXX & "%' ORDER BY ID")
    End If

    If Not (RSTMP.EOF And RSTMP.BOF) Then
        Listview_Loadval Me.lstEstimate.ListItems, RSTMP
        lstEstimate.Refresh
        lstEstimate.Enabled = True
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function GetTaym()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim x                                              As Integer
    Dim cnt                                            As Integer
    cnt = 0
    Set RSTMP = gconDMIS.Execute("Select PromiseDate From CSMS_RepairOrder Where RO_no = '" & txtEstimateno.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        For x = 1 To Len(RSTMP!PromiseDate)
            If Mid(RSTMP!PromiseDate, x, 1) = "/" Then cnt = cnt + 1
            If cnt = 2 Then
                GetTaym = Mid(RSTMP!PromiseDate, x + 6, Len(RSTMP!PromiseDate) - x)
                Exit For
            End If
        Next
    End If

    Set RSTMP = Nothing
End Function

Function CheckIfThereAPMS(VRO As String) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT JOBTYPE FROM CSMS_RO_DET WHERE REP_OR = '" & VRO & "' AND LIVIL = '1' AND JOBTYPE = 'PMS'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfThereAPMS = True
    Else
        CheckIfThereAPMS = False
    End If

    Set RSTMP = Nothing
End Function

Private Sub cmdPrint_Click()
    If txtEstimateno.Text = "" Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    rptRepairOrder.WindowShowPrintSetupBtn = True
    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptRepairOrder.WindowTitle = "Estimate Details"
    
    If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HPC" Then
        rptRepairOrder.Formulas(3) = "TAYM = '" & GetTaym & "'"
    End If

    If COMPANY_CODE = "HAS" Then
        If CheckIfThereAPMS(txtEstimateno) = True Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "PrintEstimate.rpt", "{repor.rep_or} = '" & txtEstimateno & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "PrintEstimate_NOPMS.rpt", "{repor.rep_or} = '" & txtEstimateno & "'", CSMS_REPORT_CONNECTION, 1
        End If
    Else
        If COMPANY_CODE = "HAI" Then
            Dim RSTMP                                  As New ADODB.Recordset
            Dim FJOB                                   As String
            Dim SJOB                                   As String
            Dim TJOB                                   As String
            Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & txtPlanteNo & "' AND TRANSTYPE = 'R' AND DTE_RECD < '" & Date & "' ORDER BY DTE_RECD ASC ")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                If Not RSTMP.BOF Then
                    RSTMP.MoveFirst
                    FJOB = Null2String(RSTMP!rep_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
                    RSTMP.MoveNext

                    If Not RSTMP.EOF Then
                        SJOB = Null2String(RSTMP!rep_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
                        RSTMP.MoveNext

                        If Not RSTMP.EOF Then
                            TJOB = Null2String(RSTMP!rep_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
                        End If
                    End If
                End If
            End If
            Set RSTMP = Nothing
            rptRepairOrder.Formulas(0) = "RO1 = '" & FJOB & "'"
            rptRepairOrder.Formulas(1) = "RO2 = '" & SJOB & "'"
            rptRepairOrder.Formulas(2) = "RO3 = '" & TJOB & "'"
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "PrintEstimate.rpt", "{repor.rep_or} = '" & txtEstimateno & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "PrintEstimate.rpt", "{repor.rep_or} = '" & txtEstimateno & "'", CSMS_REPORT_CONNECTION, 1
        End If
    End If

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "ESTIMATE", "", FindTransactionID(N2Str2Null(txtEstimateno), "REP_OR", "CSMS_REPOR"), "", "ESTIMATE NO : " & txtEstimateno & " - VIEW ESTIMATE DETAILS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Screen.MousePointer = 0
End Sub

Private Sub cmdProces_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    
    If txtEstimateno.Text = "" Then
        MsgBox "Please selecet an Estimate first", vbInformation, "Info"
        Exit Sub
    End If
    
    Set RSTMP = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE REP_OR = '" & txtROno.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        MsgBox "Repair Order No Already Exist", vbExclamation, "CSMS"
        txtROno.SetFocus
        Exit Sub
    End If

    If MsgBox("Upload this Estimate to Repair Order, Are you Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("UPDATE CSMS_ESTHD SET STATUS = 'Y' " & _
        ", DATE_UPLOAD = " & N2Str2Null(Date) & _
        ", REP_OR = '" & txtROno & _
        "' WHERE ESTIMATENO = '" & txtEstimateno & "'")
    
    gconDMIS.Execute ("UPDATE CSMS_ESTDETAILS SET " & _
        " REP_OR = '" & txtROno & _
        "' WHERE ESTIMATENO = '" & txtEstimateno & "'")
    
    SQL_STATEMENT = "update CSMS_Repor set" & _
        " REP_OR = '" & txtROno & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtEstimateno & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UP", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtEstimateno), "estimateno", "CSMS_REPOR"), "", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_Ro_Det set" & _
        " REP_OR = '" & txtROno & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtEstimateno & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtEstimateno), "estimateno", "CSMS_REPOR"), "", "EST NO: " & txtEstimateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
        " RO_No = '" & txtROno & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtEstimateno & "'"

    gconDMIS.Execute "update CSMS_PMS_Job_Det set" & _
        " REP_OR = '" & txtROno & "'," & _
        " transtype = 'R'" & _
        " where estimateno = '" & txtEstimateno & "'"
    
    gconDMIS.Execute ("DELETE FROM CSMS_RO_dET WHERE LIVIL <> 1")
    
    MessagePop InfoFriend, "Estimate Succesfully uploaded to Repair Order", "Estimate Uploaded", 1000
    
    cmdCancel.Value = True
End Sub

Private Sub cmdEdit_Click()
    If Not lstEstimate.ListItems.Count = 0 Then
        frmCSMSEdit.Option2.Value = True
        frmCSMSEdit.Show 1
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call txtKeyword_Change
    Call GenerateNewRONO
End Sub

Private Sub lstEstimate_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE ESTIMATENO = '" & Item.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtEstimateno.Text = Null2String(RSTMP!EstimateNo)
        txtCustomer.Text = Null2String(RSTMP!NIYM)
        txtPlanteNo.Text = Null2String(RSTMP!PLATE_NO)
        txtModel.Text = Null2String(RSTMP!MODEL)
        txtKM.Text = Null2String(RSTMP!km_rdg)
        TXTWrite.Text = Null2String(FindSAName(RSTMP!recd_by))
        txtdate.Text = Null2String(RSTMP!DTE_RECD)
    End If

    Set RSTMP = Nothing
End Sub

Private Sub Timer1_Timer()
    If A.ForeColor = &HC0& Then
        A.ForeColor = &HC0C0&
    Else
        A.ForeColor = &HC0&
    End If
End Sub

Private Sub txtKeyword_Change()
    If txtKeyword.Text = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(txtKeyword)
    End If
    '    Set rsUpload = New ADODB.Recordset
    '    lstEstimate.Enabled = False
    '    lstEstimate.Sorted = False: lstEstimate.ListItems.Clear'''

    '    If Option1.Value = True Then
    '        Set rsUpload = gconDMIS.Execute("select ESTIMATENO,niym,id CSMS_repor where ro_no is null and estimateno like '" & txtKeyword & "%' order by estimateno asc")
    '    ElseIf Option2.Value = True Then
    '        Set rsUpload = gconDMIS.Execute("select ESTIMATENO,Customer,plate_no,Model,Status,acct_no from CSMS_vw_REPAIRORDER where ro_no is null and Customer like '" & txtKeyword & "%' order by Customer asc")
    '    End If
    '
    '    If Not rsUpload.EOF And Not rsUpload.BOF Then
    '        Call Listview_Loadval(Me.lstEstimate.ListItems, rsUpload)
    '        lstEstimate.Enabled = True
    '    End If
End Sub
