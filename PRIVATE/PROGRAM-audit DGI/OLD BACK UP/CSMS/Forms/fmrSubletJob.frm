VERSION 5.00
Begin VB.Form frmCSMS_SubletJob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sublet Repair Data Entry"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "fmrSubletJob.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   8145
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   5145
   End
   Begin VB.TextBox txtROno 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "Job Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   11
      Top             =   690
      Width           =   7965
      Begin VB.ComboBox cboBP_TYPE 
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
         ItemData        =   "fmrSubletJob.frx":058A
         Left            =   5010
         List            =   "fmrSubletJob.frx":0594
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1620
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.ComboBox cboBPorGJ 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "fmrSubletJob.frx":05A6
         Left            =   6060
         List            =   "fmrSubletJob.frx":05B0
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   1125
      End
      Begin VB.ComboBox cboSubletCategory 
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
         ItemData        =   "fmrSubletJob.frx":05BC
         Left            =   1200
         List            =   "fmrSubletJob.frx":05C9
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3645
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
         Height          =   855
         Left            =   7050
         MouseIcon       =   "fmrSubletJob.frx":05FB
         MousePointer    =   99  'Custom
         Picture         =   "fmrSubletJob.frx":074D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   4500
         Width           =   855
      End
      Begin VB.TextBox txtCompAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2820
         Width           =   1755
      End
      Begin VB.TextBox txtContracAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   2100
         TabIndex        =   6
         Top             =   1650
         Width           =   1755
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   810
         Width           =   4815
      End
      Begin VB.ComboBox cboJobChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "fmrSubletJob.frx":0A8B
         Left            =   5010
         List            =   "fmrSubletJob.frx":0A98
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   765
      End
      Begin VB.TextBox txtOPCODE 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   1755
      End
      Begin VB.TextBox txtSubletAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   2100
         TabIndex        =   5
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Frame Frame2 
         Caption         =   "Suggested Jobs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   90
         TabIndex        =   19
         Top             =   2100
         Width           =   7815
         Begin VB.TextBox txtNote 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   60
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.TextBox txtPaintMaterials 
         Alignment       =   1  'Right Justify
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
         Left            =   5730
         TabIndex        =   23
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6150
         MouseIcon       =   "fmrSubletJob.frx":0AA5
         MousePointer    =   99  'Custom
         Picture         =   "fmrSubletJob.frx":0BF7
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save Entry"
         Top             =   4500
         Width           =   915
      End
      Begin VB.Label lbltechcode 
         Caption         =   "lbltechcode"
         Height          =   225
         Left            =   2040
         TabIndex        =   30
         Top             =   5070
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "BP Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4170
         TabIndex        =   28
         Top             =   1710
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Classification"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   420
         Width           =   1995
      End
      Begin VB.Label LINE_NO 
         Caption         =   "LINE_NO"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   5070
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label labDET 
         Caption         =   "labDet"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   4830
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblAddorEdit 
         Caption         =   "lblAddorEdit"
         Height          =   225
         Left            =   180
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "CONTRACTOR AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   1770
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "COMPANY AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8130
         TabIndex        =   17
         Top             =   2610
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   30
         TabIndex        =   14
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3990
         TabIndex        =   13
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label labDetCost 
         Caption         =   "SUBLET LABOR AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   1380
         Width           =   1995
      End
      Begin VB.Label lblMatAmount 
         Caption         =   "PAINT MAT. AMOUNT"
         Height          =   435
         Left            =   3990
         TabIndex        =   22
         Top             =   990
         Width           =   2205
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Customer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmCSMS_SubletJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJOBS                                             As ADODB.Recordset
Dim rsROJOBS                                           As ADODB.Recordset
Dim JOB_GROUP                                          As String
Dim Sublet_TOTAL_AMT                                   As Double
Dim Sublet_TOTAl_VAT                                   As Double
Dim Sublet_NET_AMT                                     As Double
Dim VPo_No                                             As String
Dim vLIVIL                                             As String
Dim vDetAmntWoVat                                      As Double
Dim vTaxVal                                            As Double
Dim vtxtSubletAmount                                   As Double
Dim vIFvat                                             As String


Function GetJobLineNo(xxx As String)
    Dim rsGetLN                                        As New ADODB.Recordset
    Dim tempMAX_LINE_NO                                As String
    Dim vFinalResult                                   As String

    GetJobLineNo = "80"

    'CSMS_PO_DT
    Set rsGetLN = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO ,REP_OR from CSMS_PO_DT where [REP_OR] = '" & xxx & "' and LIVIL =" & vLIVIL & " order by MAX_LINE_NO desc")
    If Not rsGetLN.EOF And Not rsGetLN.BOF Then
        tempMAX_LINE_NO = NumericVal(rsGetLN!MAX_LINE_NO)
    Else
        tempMAX_LINE_NO = NumericVal(GetJobLineNo)
    End If
    Set rsGetLN = Nothing

    Dim vResult                                        As String
    Dim rsRoDet_lineNO                                 As New ADODB.Recordset
    Dim tempMAX_LINE_NO2                               As String
    Set rsRoDet_lineNO = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO2 ,REP_OR from CSMS_RO_DET where [REP_OR] = '" & xxx & "' and LIVIL =" & vLIVIL & " order by MAX_LINE_NO2 desc")
    If Not rsRoDet_lineNO.EOF And Not rsRoDet_lineNO.BOF Then
        tempMAX_LINE_NO2 = NumericVal(rsRoDet_lineNO!MAX_LINE_NO2)
    Else
        tempMAX_LINE_NO2 = NumericVal(GetJobLineNo)
    End If
    Set rsRoDet_lineNO = Nothing

    Dim rsRC_Line_no                                   As New ADODB.Recordset
    Dim tempMAX_LINE_NO3                               As String
    Set rsRC_Line_no = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO3 ,REP_OR from CSMS_PO_RC_DT where [REP_OR] = '" & xxx & "' and LIVIL =" & vLIVIL & " order by MAX_LINE_NO3 desc")
    If Not rsRC_Line_no.EOF And Not rsRC_Line_no.BOF Then
        tempMAX_LINE_NO3 = NumericVal(rsRC_Line_no!MAX_LINE_NO3)
    Else
        tempMAX_LINE_NO3 = NumericVal(GetJobLineNo)
    End If

    If NumericVal(tempMAX_LINE_NO) > NumericVal(tempMAX_LINE_NO2) Then
        vResult = NumericVal(tempMAX_LINE_NO)
    ElseIf NumericVal(tempMAX_LINE_NO) = NumericVal(tempMAX_LINE_NO2) Then
        vResult = NumericVal(tempMAX_LINE_NO)
    Else
        vResult = NumericVal(tempMAX_LINE_NO2)
    End If

    If NumericVal(vResult) > NumericVal(tempMAX_LINE_NO3) Then
        GetJobLineNo = Format(NumericVal(vResult) + 1, "00")
    ElseIf NumericVal(vResult) = NumericVal(tempMAX_LINE_NO3) Then
        GetJobLineNo = Format(NumericVal(vResult) + 1, "00")
    Else
        GetJobLineNo = Format(NumericVal(tempMAX_LINE_NO3) + 1, "00")
    End If

    Set rsRC_Line_no = Nothing
End Function

Function SetContractorCode(xxx As String) As String
    Dim rsContractorAdd                                As New ADODB.Recordset
    Dim rsVENDOR                                       As New ADODB.Recordset
    
    Set rsContractorAdd = gconDMIS.Execute("Select * from CSMS_Contractor Where CompanyName = '" & xxx & "'")
    If Not rsContractorAdd.EOF And Not rsContractorAdd.BOF Then
        SetContractorCode = Null2String(rsContractorAdd!Code)
    Else
        Set rsVENDOR = gconDMIS.Execute("select CODE FROM ALL_VENDOR_TABLE WHERE NAMEOFVENDOR = '" & xxx & "'")
        If Not (rsVENDOR.BOF And rsVENDOR.EOF) Then
            SetContractorCode = Null2String(rsVENDOR!Code)
        End If
    End If
    Set rsContractorAdd = Nothing
End Function

Function PO_GetID(xxx As Variant) As Variant
    Dim JUN                                            As New ADODB.Recordset
    Dim idCaption                                      As String

    Set JUN = gconDMIS.Execute("Select ID from CSMS_PO_hd where Po_no = '" & xxx & "'")
    If Not JUN.EOF And Not JUN.EOF Then
        idCaption = Null2String(JUN!ID)
        frmCSMS_PurchaseOrder.passID (idCaption)
    End If
    Set JUN = Nothing
End Function

Function RC_GetID(xxx As Variant) As Variant
    Dim MINNIE                                         As ADODB.Recordset
    Dim idCaption                                      As String

    Set MINNIE = gconDMIS.Execute("Select ID from CSMS_PO_Rc_hd where RC_NO = '" & xxx & "'")
    If Not MINNIE.EOF And Not MINNIE.EOF Then
        idCaption = Null2String(MINNIE!ID)
        frmCSMS_ReceivingEntry.passID (idCaption)
    End If
    Set MINNIE = Nothing
End Function

Sub initmamvars()
    txtSubletAmount.Text = Format(NumericVal(txtSubletAmount), MAXIMUM_DIGIT)
    txtCompAmount.Text = Format(NumericVal(txtCompAmount), MAXIMUM_DIGIT)
    txtContracAmount.Text = Format(NumericVal(txtContracAmount), MAXIMUM_DIGIT)
    txtPaintMaterials.Text = Format(NumericVal(txtPaintMaterials), MAXIMUM_DIGIT)
End Sub

Sub ComputeTotalCost()
    Dim rsComputeTotalCost                             As New ADODB.Recordset
    Dim VPo_No                                         As String
    Dim VRC_NUM                                        As String

    Sublet_TOTAL_AMT = 0: Sublet_TOTAl_VAT = 0: Sublet_NET_AMT = 0
    If lblAddorEdit.Caption = "ADD" Or lblAddorEdit.Caption = "EDIT" Then
        VPo_No = frmCSMS_PurchaseOrder.txtPoNumber.Text

        Set rsComputeTotalCost = gconDMIS.Execute("Select DETAMT,TAXVAL,DET_AMT from CSMS_PO_dt where PO_No ='" & VPo_No & "'")
        If Not rsComputeTotalCost.EOF And Not rsComputeTotalCost.BOF Then
            Do While Not rsComputeTotalCost.EOF
                Sublet_TOTAL_AMT = N2Str2Zero(rsComputeTotalCost!DETAMT) + N2Str2Zero(Sublet_TOTAL_AMT)
                Sublet_TOTAl_VAT = N2Str2Zero(rsComputeTotalCost!TAXVAL) + N2Str2Zero(Sublet_TOTAl_VAT)
                Sublet_NET_AMT = N2Str2Zero(rsComputeTotalCost!DET_AMT) + N2Str2Zero(Sublet_NET_AMT)
                rsComputeTotalCost.MoveNext
            Loop
        End If
    Else
        VRC_NUM = frmCSMS_ReceivingEntry.txtRcNumber.Text
        Set rsComputeTotalCost = gconDMIS.Execute("Select DETAMT,TAXVAL,DET_AMT from CSMS_PO_RC_DT where RC_NO ='" & VRC_NUM & "'")
        If Not rsComputeTotalCost.EOF And Not rsComputeTotalCost.BOF Then
            Do While Not rsComputeTotalCost.EOF
                Sublet_TOTAL_AMT = N2Str2Zero(rsComputeTotalCost!DETAMT) + N2Str2Zero(Sublet_TOTAL_AMT)
                Sublet_TOTAl_VAT = N2Str2Zero(rsComputeTotalCost!TAXVAL) + N2Str2Zero(Sublet_TOTAl_VAT)
                Sublet_NET_AMT = N2Str2Zero(rsComputeTotalCost!DET_AMT) + N2Str2Zero(Sublet_NET_AMT)
                rsComputeTotalCost.MoveNext
            Loop
        End If
    End If
End Sub

Sub AdditionalJobFromReceiving()
    Dim rRC_NO                                         As String
    Dim rPo_No                                         As String
    Dim rRep_or                                        As String
    Dim rJOBTYPE                                       As String
    Dim rLIVIL                                         As String
    Dim rLINE_NO                                       As String
    Dim rDETAMT                                        As Double
    Dim rDETCDE                                        As String
    Dim rDETDSC                                        As String
    Dim rTECHNICIAN                                    As String
    Dim rWCODE                                         As String
    Dim rTAXRATE                                       As Double
    Dim rTAXVAL                                        As Double
    Dim rSTATUS                                        As String
    Dim rDETAIL                                        As String
    Dim rDET_AMT                                       As Double
    Dim rUSERCODE                                      As String
    Dim rSAVEDATE                                      As String
    Dim rTECHCODE                                      As String
    Dim rCOMPAMOUNT                                    As Double
    Dim rCONTRACAMOUNT                                 As Double
    Dim rDONE                                          As String
    Dim rROTYPE                                        As String

    rDETAMT = 0: rTAXRATE = 0: rTAXVAL = 0: rDET_AMT = 0

    If cboSubletCategory.Text = "" Then
        MsgBox "Pls...Select a Sublet Job Category", vbInformation, "INFORMATION"
        txtOPCODE.Text = ""
        txtJobDesc.Text = ""
        cboSubletCategory.SetFocus
        Exit Sub
    End If

    If cboBPorGJ.Text = "" Then
        MsgBox "Pls...Select Job Classification", vbInformation, "INFORMATION"
        cboBPorGJ.SetFocus
        Exit Sub
    End If

    If txtNote.Text = "" Then
        MsgBox "Pls...Type the suggested Job", vbInformation, "INFORMATION"
        txtNote.SetFocus
        Exit Sub
    End If

    rRC_NO = N2Str2Null(frmCSMS_ReceivingEntry.txtRcNumber.Text)
    rPo_No = N2Str2Null(frmCSMS_ReceivingEntry.cboPoNumber.Text)
    rRep_or = N2Str2Null(txtROno.Text)
    rROTYPE = "'SR'"
    rJOBTYPE = N2Str2Null(cboBPorGJ.Text)

    If lblAddorEdit.Caption = "RADD" Then
        rLINE_NO = N2Str2Null(GetJobLineNo(txtROno.Text))
    Else
        rLINE_NO = Format(N2Str2Null(LINE_NO.Caption), "00")
    End If
    If cboJobChargeTo = "C" Or cboJobChargeTo = "S" Then
        rDETAMT = NumericVal(txtSubletAmount)
    Else
        rDETAMT = NumericVal(txtSubletAmount) / 1.12
    End If
    rDETCDE = N2Str2Null(txtOPCODE.Text)
    rDETDSC = N2Str2Null(txtJobDesc.Text)
    rTECHNICIAN = N2Str2Null(frmCSMS_ReceivingEntry.cboContractor.Text)
    rWCODE = N2Str2Null(cboJobChargeTo.Text)
    If cboJobChargeTo = "C" Or cboJobChargeTo = "S" Then
        rTAXRATE = 0
    Else
        rTAXRATE = (VAT_RATE / 100)
    End If
    If cboJobChargeTo = "C" Or cboJobChargeTo = "S" Then
        rTAXVAL = 0
    Else
        rTAXVAL = Round(((txtSubletAmount) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    End If
    rSTATUS = "NULL"
    rDETAIL = N2Str2Null(txtNote.Text)
    rDET_AMT = NumericVal(txtSubletAmount)
    rUSERCODE = "" & N2Str2Null(LOGCODE) & ""
    rSAVEDATE = DateValue(Now) & " " & TimeValue(Now)
    rTECHCODE = N2Str2Null(frmCSMS_ReceivingEntry.txtReceive.Text)
    rCONTRACAMOUNT = NumericVal(txtContracAmount.Text)
    rCOMPAMOUNT = NumericVal(txtCompAmount.Text)
    rDONE = "'Y'"

    If lblAddorEdit.Caption = "RADD" Then
        If MsgBox("Are you sure you want to add this Job", vbYesNo, vbInformation) = vbYes Then
            gconDMIS.Execute "Insert into CSMS_PO_RC_Dt " & _
                "(RC_NO,Po_No,Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETAMT,DETCDE,DETDSC,TECHNICIAN,WCODE,TAXRATE,TAXVAL,STATUS,DETAIL,DET_AMT,USERCODE,SAVEDATE,TECHCODE,CONTRACTAMOUNT,COMPAMOUNT,DONE)" & _
                "VALUES(" & rRC_NO & _
                "," & rPo_No & _
                "," & rRep_or & _
                "," & rROTYPE & _
                "," & rJOBTYPE & _
                "," & vLIVIL & _
                "," & rLINE_NO & _
                "," & rDETAMT & _
                "," & rDETCDE & _
                "," & rDETDSC & _
                "," & rTECHNICIAN & _
                "," & rWCODE & _
                "," & rTAXRATE & _
                "," & rTAXVAL & _
                "," & rSTATUS & _
                "," & rDETAIL & _
                "," & rDET_AMT & _
                "," & rUSERCODE & _
                ",'" & rSAVEDATE & _
                "'," & rTECHCODE & _
                "," & rCONTRACAMOUNT & _
                "," & rCOMPAMOUNT & _
                "," & rDONE & ")"

            'insert to CSMS_RO_DET
            gconDMIS.Execute "Insert into CSMS_RO_DET" & _
                "(Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,TECHNICIAN,DETAMT,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,STATUS,USERCDE,SAVEDATE,DONE,SUBPOCODE,TECHCODE)" & _
                "VALUES(" & rRep_or & _
                "," & rROTYPE & _
                "," & rJOBTYPE & _
                "," & vLIVIL & _
                "," & rLINE_NO & _
                "," & rDETCDE & _
                "," & rDETDSC & _
                "," & rTECHNICIAN & _
                "," & rDETAMT & _
                "," & rWCODE & _
                "," & VAT_RATE & _
                "," & rTAXVAL & _
                "," & rDETAIL & _
                "," & rDET_AMT & _
                "," & rDONE & _
                "," & rUSERCODE & _
                ",'" & rSAVEDATE & _
                "'," & rDONE & _
                "," & rPo_No & _
                "," & rTECHCODE & ")"

            Call ShowSuccessFullyAdded

            'insert to CSMS_PO_DT <--------------06-18-2008-----------------------------
            gconDMIS.Execute "Insert into CSMS_PO_DT" & _
                "(Po_No,Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETAMT,DETCDE,DETDSC,TECHNICIAN,wCode,TAXRATE,TAXVAL,Status,DETAIL,DET_AMT,USERCODE,SAVEDATE,TechCode,CONTRACTAMOUNT,COMPAMOUNT,DONE)" & _
                "VALUES(" & rPo_No & _
                "," & rRep_or & _
                "," & rROTYPE & _
                "," & rJOBTYPE & _
                "," & vLIVIL & _
                "," & rLINE_NO & _
                "," & rDETAMT & _
                "," & rDETCDE & _
                "," & rDETDSC & _
                "," & rTECHNICIAN & _
                "," & rWCODE & _
                "," & rTAXRATE & _
                "," & rTAXVAL & _
                "," & rSTATUS & _
                "," & rDETAIL & _
                "," & rDET_AMT & _
                "," & rUSERCODE & _
                ",'" & rSAVEDATE & _
                "'," & rTECHCODE & _
                "," & rCONTRACAMOUNT & _
                "," & rCOMPAMOUNT & _
                "," & rDONE & ")"
            '------------------------------------------------------------------->
        Else
            Exit Sub
        End If

        Call ComputeTotalCost

        gconDMIS.Execute "Update CSMS_PO_RC_HD set " & _
            " SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & _
            ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & _
            ", SUBLET_TOTAL_NET_AMT =" & Sublet_NET_AMT & _
            " WHERE RC_NO = '" & frmCSMS_ReceivingEntry.txtRcNumber.Text & "'"
        
        Call RC_GetID(frmCSMS_ReceivingEntry.txtRcNumber.Text)
        frmCSMS_ReceivingEntry.labDET.Caption = ""
        'update CSMS_PO_HD <---------------------06-18-2008-----------------------------------
        gconDMIS.Execute "Update CSMS_PO_HD set " & _
            " SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & _
            ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & _
            ", SUBLET_TOTAL_NET_AMT = " & Sublet_NET_AMT & _
            " WHERE PO_NO = " & rPo_No & ""
        '------------------------------------------------------------------------------------->

        '            '<----------check the status of the repair order
        '            Dim rsRO_DET As ADODB.Recordset
        '            Set rsRO_DET = New ADODB.Recordset
        '                 Set rsRO_DET = gconDMIS.Execute("Select STATUS from CSMS_RO_DET WHERE LIVIL = '1' AND REP_OR = '" & txtROno & "' AND (DONE = 'N' OR DONE IS NULL OR DONE ='W')")
        '
        '                 If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        '                     gconDMIS.Execute "Update CSMS_RepairOrder set JSTATUS ='W', STATUS ='Working' where Ro_No ='" & txtROno & "'"
        '                 Else
        '                     gconDMIS.Execute "Update CSMS_RepairOrder set dateFinish = '" & LOGDATE & "', JSTATUS ='F', STATUS ='Finish Job' where RO_NO ='" & txtROno & "'"
        '                 End If
        ''            Set rsRO_DET = Nothing
        '            '----------------->
        Call RC_GetID(frmCSMS_ReceivingEntry.txtRcNumber.Text)
        'PO_GetID (frmCSMS_PurchaseOrder.txtPoNumber.Text)

        Unload Me
    Else
        If MsgBox("Are you sure you want to save this Job", vbYesNo + vbInformation, "INFORMATION") = vbYes Then
            gconDMIS.Execute "Update CSMS_PO_RC_DT set " & _
                "RC_NO = " & rRC_NO & "," & _
                "Po_No = " & rPo_No & "," & _
                "Rep_or = " & rRep_or & "," & _
                "JOBTYPE = " & rJOBTYPE & "," & _
                "LIVIL = " & vLIVIL & "," & _
                "LINE_NO= " & rLINE_NO & "," & _
                "DETAMT = " & rDETAMT & "," & _
                "DETCDE=" & rDETCDE & "," & _
                "DETDSC =" & rDETDSC & "," & _
                "TECHNICIAN = " & rTECHNICIAN & "," & _
                "WCODE = " & rWCODE & "," & _
                "TAXRATE =" & rTAXRATE & "," & _
                "TAXVAL = " & rTAXVAL & "," & _
                "STATUS = " & rSTATUS & "," & _
                "DETAIL = " & rDETAIL & "," & _
                "DET_AMT = " & rDET_AMT & "," & _
                "USERCODE = " & rUSERCODE & "," & _
                "SAVEDATE = '" & rSAVEDATE & "'," & _
                "TECHCODE = " & rTECHCODE & "," & _
                "CONTRACTAMOUNT = " & rCONTRACAMOUNT & "," & _
                "COMPAMOUNT = " & rCOMPAMOUNT & "," & _
                "DONE = " & rDONE & " " & _
                "Where ID =" & labDET

            'Update the Final amount in the CSMS_RO_DET......
            gconDMIS.Execute "Update CSMS_RO_DET set " & _
                " JOBTYPE = " & rJOBTYPE & _
                ", DETAIL = " & rDETAIL & _
                ", LIVIL = " & vLIVIL & _
                ", DETCOST = " & rCONTRACAMOUNT & _
                ", DETAMT = " & rDETAMT & _
                ", DET_AMT = " & rDET_AMT & _
                ", WCODE = " & rWCODE & _
                ", TAXVAL = " & rTAXVAL & _
                " where Rep_or = " & rRep_or & _
                " and LINE_NO = " & rLINE_NO & _
                " and  LIVIL = " & vLIVIL & ""

            'update the final amount in the csms_po_dt
            gconDMIS.Execute "Update CSMS_PO_DT set " & _
                " JOBTYPE = " & rJOBTYPE & _
                ", DETAIL = " & rDETAIL & _
                ", LIVIL = " & vLIVIL & _
                ", CONTRACTAMOUNT = " & rCONTRACAMOUNT & _
                ", TAXVAL = " & rTAXVAL & _
                ", TAXRATE = " & rTAXRATE & _
                ", WCODE = '" & cboJobChargeTo & _
                "', COMPAMOUNT = " & rCOMPAMOUNT & _
                " WHERE PO_NO = " & rPo_No & _
                " AND LIVIL = " & vLIVIL & _
                " AND LINE_NO = " & rLINE_NO & ""

            txtSubletAmount.Enabled = True
            Call ShowSuccessFullyUpdated
        Else
            Exit Sub
        End If

        Call ComputeTotalCost
        gconDMIS.Execute "Update CSMS_PO_RC_HD set SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & ", SUBLET_TOTAL_NET_AMT =" & Sublet_NET_AMT & " WHERE RC_NO ='" & frmCSMS_ReceivingEntry.txtRcNumber.Text & "'"
        gconDMIS.Execute "Update CSMS_PO_HD set SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & ", SUBLET_TOTAL_NET_AMT =" & Sublet_NET_AMT & " WHERE PO_NO ='" & frmCSMS_ReceivingEntry.cboPoNumber.Text & "'"
        Call RC_GetID(frmCSMS_ReceivingEntry.txtRcNumber.Text)
        frmCSMS_ReceivingEntry.labDET.Caption = ""

        Unload Me
    End If
End Sub

Private Sub cboBPorGJ_Change()
    If UCase(cboSubletCategory.Text) = "SUBLET LABOR" Then
        If UCase(cboBPorGJ.Text) = "GJ" Then
            cboBP_TYPE.Visible = False
            Label4.Visible = False
        Else
            cboBP_TYPE.Visible = True
            Label4.Visible = True
        End If
    End If
End Sub

Private Sub cboBPorGJ_Click()
    If UCase(cboSubletCategory.Text) = "SUBLET LABOR" Then
        If UCase(cboBPorGJ.Text) = "GJ" Then
            cboBP_TYPE.Visible = False
            Label4.Visible = False
        Else
            cboBP_TYPE.Visible = True
            Label4.Visible = True
        End If
    End If
End Sub

Private Sub cboSubletCategory_Click()
    If cboSubletCategory.Text = "SUBLET LABOR" Then
        txtOPCODE.Text = "SRLABOR"
        txtJobDesc.Text = "SUBLET FOR LABOR:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'1'"
        'txtJobDesc.SetFocus
    ElseIf cboSubletCategory.Text = "SUBLET PARTS" Then
        txtOPCODE.Text = "SRPARTS"
        txtJobDesc.Text = "SUBLET FOR PARTS:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'2'"
        'txtJobDesc.SetFocus
    Else
        txtOPCODE.Text = "SRMATERIALS"
        txtJobDesc.Text = "SUBLET FOR MATERIALS:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'3'"
        'txtJobDesc.SetFocus
    End If
End Sub

Private Sub cboSubletCategory_LostFocus()
    If cboSubletCategory.Text = "SUBLET LABOR" Then
        txtOPCODE.Text = "SRLABOR"
        txtJobDesc.Text = "SUBLET FOR LABOR:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'1'"
        'txtJobDesc.SetFocus
    ElseIf cboSubletCategory.Text = "SUBLET PARTS" Then
        txtOPCODE.Text = "SRPARTS"
        txtJobDesc.Text = "SUBLET FOR PARTS:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'2'"
        'txtJobDesc.SetFocus
    Else
        'txtJobDesc.SetFocus
        txtOPCODE.Text = "SRMATERIALS"
        txtJobDesc.Text = "SUBLET FOR MATERIALS:"
        'txtNote.Text = "SUBLET FOR LABOR:"
        vLIVIL = "'3'"
        'txtJobDesc.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    If lblAddorEdit.Caption = "EDIT" Or lblAddorEdit.Caption = "REDIT" Then
        If lblAddorEdit.Caption = "EDIT" Then
            frmCSMS_PurchaseOrder.cmdCancel = True
            Unload Me
            'frmCSMS_PurchaseOrder.EnabledFrame (True)
        Else
            frmCSMS_ReceivingEntry.cmdCancel = True
            Unload Me
            'frmCSMS_PurchaseOrder.EnabledFrame (True)
        End If
    ElseIf lblAddorEdit.Caption = "ADD" Or lblAddorEdit.Caption = "RADD" Then
        Unload Me
        'frmCSMS_PurchaseOrder.EnabledFrame (True)
    End If
End Sub

Private Sub cmdCancel2_Click()
    If lblAddorEdit = "ADD" Or lblAddorEdit = "RADD" Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Call initmamvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lblAddorEdit.Caption = "RADD" Or lblAddorEdit.Caption = "REDIT" Then
        frmCSMS_ReceivingEntry.EnabledFrame (True)
        Unload Me
    Else
        frmCSMS_PurchaseOrder.EnabledFrame (True)
        Unload Me
    End If
End Sub

Private Sub txtContracAmount_Change()
    txtCompAmount.Text = Format(NumericVal(txtSubletAmount.Text), MAXIMUM_DIGIT) - Format(NumericVal(txtContracAmount.Text), MAXIMUM_DIGIT)
    txtCompAmount.Text = Format(NumericVal(txtCompAmount.Text), MAXIMUM_DIGIT)

    If NumericVal(txtContracAmount.Text) > NumericVal(txtSubletAmount.Text) Then
        MsgBox "The amount you entered is greater" & vbCrLf & "than the total sublet amount"
        txtContracAmount.Text = Left(txtContracAmount.Text, Len(txtSubletAmount) - 4)
        txtContracAmount.Text = Format(NumericVal(txtContracAmount), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtContracAmount_GotFocus()
    If txtContracAmount.Text = "0.00" Or txtContracAmount.Text = "" Then
        txtContracAmount.Text = ""
        txtCompAmount.Text = "0.00"
    Else
        txtContracAmount.Text = Format(NumericVal(txtContracAmount), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtContracAmount_LostFocus()
    If txtContracAmount.Text = "" Then
        txtContracAmount.Text = "0.00"
    Else
        txtContracAmount.Text = Format(NumericVal(txtContracAmount), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtSubletAmount_GotFocus()
    If txtSubletAmount.Text = "0.00" Then
        txtSubletAmount.Text = ""
    Else
        txtSubletAmount.Text = Format(NumericVal(txtSubletAmount), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtSubletAmount_LostFocus()
    If txtSubletAmount.Text = "" Then
        txtSubletAmount.Text = "0.00"
    Else
        txtSubletAmount.Text = Format(NumericVal(txtSubletAmount), MAXIMUM_DIGIT)
        'txtContracAmount.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If lblAddorEdit.Caption = "RADD" Then
        Call AdditionalJobFromReceiving
    ElseIf lblAddorEdit.Caption = "REDIT" Then
        Call AdditionalJobFromReceiving
    Else
        Dim rstmp                                      As New ADODB.Recordset

        If cboSubletCategory.Text = "" Then
            MsgBox "Pls... Select a Job Category.", vbInformation, "INFORMATION"
            txtOPCODE.Text = ""
            txtJobDesc.Text = ""
            cboSubletCategory.SetFocus
            Exit Sub
        End If

        If cboBPorGJ.Text = "" Then
            MsgBox "Pls...Select Job Classification", vbInformation, "INFORMATION"
            cboBPorGJ.SetFocus
            Exit Sub
        End If

        If txtNote.Text = "" Then
            MsgBox "Pls...Type the Suggested Sublet Job.", vbInformation, "INFORMATION"
            txtNote.SetFocus
            Exit Sub
        End If

        If UCase(cboSubletCategory.Text) = "SUBLET LABOR" Then
            If cboBPorGJ.Text = "BP" Then
                If cboBP_TYPE.Text = "" Then
                    ShowIsRequiredMsg "BP Type Cannot be Blank"
                    cboBP_TYPE.SetFocus
                    Exit Sub
                End If
            End If
        End If

        If NumericVal(txtSubletAmount) = 0 Then
            ShowIsRequiredMsg "Sublet amount cannot be Zero or blank"
            txtSubletAmount.SetFocus
            Exit Sub
        End If
        
        If NumericVal(txtContracAmount) = 0 Then
            ShowIsRequiredMsg "Contractor amount cannot be Zero or blank"
            txtContracAmount.SetFocus
            Exit Sub
        End If
        
        
        Dim vPONO                                      As String
        Dim vREPOR                                     As String
        Dim vJobType                                   As String
        Dim vlineNo                                    As String
        Dim vtxtOPCODE                                 As String
        Dim vtxtJobDesc                                As String
        Dim vTechnician                                As String
        Dim vcboJobChargeTo                            As String
        Dim vTaxRate                                   As Double
        Dim VStatus                                    As String
        Dim vDetail                                    As String
        Dim Vusercode                                  As String
        Dim vSavedate                                  As String
        Dim VTECHCODE                                  As String
        Dim vtxtContracAmount                          As Double
        Dim vtxtCompAmount                             As Double
        Dim vDone                                      As String
        Dim ans                                        As String
        Dim VROTYPE                                    As String
        Dim vBP_TYPE                                   As String
        Dim vtechcode1                                 As String

        If cboBP_TYPE.Visible = False Then
            vBP_TYPE = N2Str2Null("")
        Else
            If cboBP_TYPE.Text = "Major" Then
                vBP_TYPE = N2Str2Null("M")
            Else
                vBP_TYPE = N2Str2Null("N")
            End If
        End If

        vDetAmntWoVat = 0: vTaxRate = 0
        vTaxVal = 0: vtxtSubletAmount = 0:            'vtxtPaintMaterials = 0

        vPONO = N2Str2Null(frmCSMS_PurchaseOrder.txtPoNumber.Text)
        vREPOR = N2Str2Null(txtROno.Text)
        VROTYPE = "'SR'"
        vJobType = N2Str2Null(cboBPorGJ.Text)         ' this will be change to the classification
        vtechcode1 = N2Str2Null(lbltechcode.Caption)
        If lblAddorEdit = "ADD" Then
            vlineNo = N2Str2Null(GetJobLineNo(txtROno.Text))
        Else
            vlineNo = Format(N2Str2Null(LINE_NO.Caption), "00")
        End If
'        If cboJobChargeTo.Text = "S" Or cboJobChargeTo.Text = "C" Then
'            vDetAmntWoVat = NumericVal(txtSubletAmount)
'        Else
'            vDetAmntWoVat = NumericVal(txtSubletAmount) / 1.12
'        End If
        vtxtOPCODE = N2Str2Null(txtOPCODE.Text)
        vtxtJobDesc = N2Str2Null(txtJobDesc.Text)
        vTechnician = N2Str2Null(frmCSMS_PurchaseOrder.cboContractor.Text)
        vcboJobChargeTo = N2Str2Null(cboJobChargeTo.Text)
        If cboJobChargeTo.Text = "S" Or cboJobChargeTo.Text = "C" Then
            vTaxRate = 0
        Else
            vTaxRate = (VAT_RATE / 100)
        End If
        vtxtSubletAmount = NumericVal(txtSubletAmount)
'Updated By:        IEBV 08162010_0155pm
'Description:       To check if the cotractor in non vat or with vat
        vIFvat = checkifvatornonvat(vtechcode1)
        If vIFvat = "Y" Then
            vTaxVal = 0
        Else
            If cboJobChargeTo.Text = "S" Or cboJobChargeTo.Text = "C" Then
                vTaxVal = 0
            Else
                vTaxVal = Round(((vtxtSubletAmount) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
            End If
        End If
        
        If vIFvat = "Y" Then
            vDetAmntWoVat = NumericVal(txtSubletAmount)
        Else
            If cboJobChargeTo.Text = "S" Or cboJobChargeTo.Text = "C" Then
                vDetAmntWoVat = NumericVal(txtSubletAmount)
            Else
                vDetAmntWoVat = NumericVal(txtSubletAmount) / 1.12
            End If
        End If
'----------------------------------------------------------------------------------
'        If cboJobChargeTo.Text = "S" Or cboJobChargeTo.Text = "C" Then
'            vTaxVal = 0
'        Else
'            vTaxVal = Round(((vtxtSubletAmount) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
'        End If
        vDetail = N2Str2Null(txtNote.Text)
        Vusercode = "" & N2Str2Null(LOGCODE) & ""
        vSavedate = DateValue(Now) & " " & TimeValue(Now)
        VTECHCODE = N2Str2Null(SetContractorCode(frmCSMS_PurchaseOrder.cboContractor.Text))
        vtxtContracAmount = txtContracAmount
        vtxtCompAmount = txtCompAmount
        vDone = "'N'"

        If lblAddorEdit = "ADD" Then
            If MsgBox("Are you sure do you want to add this Job?", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
                SQL_STATEMENT = "Insert into CSMS_PO_dt" & _
                    "(TRANSTATUS, Po_No,Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETAMT,DETCDE,DETDSC,TECHNICIAN,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCODE,SAVEDATE,TECHCODE,CONTRACTAMOUNT,COMPAMOUNT,DONE)" & _
                    "VALUES(" & vBP_TYPE & "," & vPONO & _
                    "," & vREPOR & _
                    "," & VROTYPE & _
                    "," & vJobType & _
                    "," & vLIVIL & _
                    "," & vlineNo & _
                    "," & vDetAmntWoVat & _
                    "," & vtxtOPCODE & _
                    "," & vtxtJobDesc & _
                    "," & vTechnician & _
                    "," & vcboJobChargeTo & _
                    "," & vTaxRate & _
                    "," & vTaxVal & _
                    "," & vDetail & _
                    "," & vtxtSubletAmount & _
                    "," & Vusercode & _
                    ",'" & vSavedate & _
                    "'," & VTECHCODE & _
                    "," & vtxtContracAmount & _
                    "," & vtxtCompAmount & _
                    "," & vDone & ")"
                gconDMIS.Execute SQL_STATEMENT
                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("AA", "SUBLET PURCHASE", SQL_STATEMENT, FindTransactionID(vPONO, "PO_NO", "CSMS_PO_HD"), "", "CODE : " & Null2String(vtxtOPCODE), vLIVIL, "")
                'NEW LOG AUDIT-----------------------------------------------------

                Call ShowSuccessFullyAdded
            Else
                Exit Sub
            End If

            Call ComputeTotalCost
            SQL_STATEMENT = "Update CSMS_PO_hd set SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & ", SUBLET_TOTAL_NET_AMT =" & Sublet_NET_AMT & " WHERE Po_No ='" & frmCSMS_PurchaseOrder.txtPoNumber.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "SUBLET PURCHASE", SQL_STATEMENT, FindTransactionID(vPONO, "PO_NO", "CSMS_PO_HD"), "", "PO NO: " & Null2String(vPONO) & " - ADD DETAILS", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            Call PO_GetID(frmCSMS_PurchaseOrder.txtPoNumber.Text)
            frmCSMS_PurchaseOrder.labDET.Caption = ""
            Unload Me
        Else
            'ans = MsgBox("Are you sure do you want to Edit this Job?", vbQuestion + vbYesNo)
            If MsgBox("Are you sure do you want to Edit this Job?", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
                SQL_STATEMENT = "Update CSMS_Po_dt set " & _
                    "TRANSTATUS = " & vBP_TYPE & ",Po_No = " & vPONO & "," & _
                    "Rep_or =" & vREPOR & "," & _
                    "ROTYPE =" & VROTYPE & "," & _
                    "JOBTYPE =" & vJobType & "," & _
                    "LIVIL = " & vLIVIL & "," & _
                    "LINE_NO = " & vlineNo & "," & _
                    "DETAMT = " & vDetAmntWoVat & "," & _
                    "DETCDE = " & vtxtOPCODE & "," & _
                    "DETDSC = " & vtxtJobDesc & "," & _
                    "TECHNICIAN = " & vTechnician & "," & _
                    "WCODE = " & vcboJobChargeTo & "," & _
                    "TAXRATE =" & vTaxRate & "," & _
                    "TAXVAL =" & vTaxVal & "," & _
                    "DETAIL = " & vDetail & "," & _
                    "DET_AMT =" & vtxtSubletAmount & "," & _
                    "USERCODE = " & Vusercode & "," & _
                    "SAVEDATE = '" & vSavedate & "'," & _
                    "TECHCODE = " & VTECHCODE & "," & _
                    "CONTRACTAMOUNT = " & vtxtContracAmount & "," & _
                    "COMPAMOUNT = " & vtxtCompAmount & "," & _
                    "DONE = " & vDone & "" & _
                    "where ID = " & labDET.Caption
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("EE", "SUBLET PURCHASE", SQL_STATEMENT, FindTransactionID(vPONO, "PO_NO", "CSMS_PO_HD"), "", "CODE : " & Null2String(vtxtOPCODE), vLIVIL, labDET)
                'NEW LOG AUDIT-----------------------------------------------------
                Call ShowSuccessFullyUpdated
            Else
                Exit Sub
            End If
            Call ComputeTotalCost

            SQL_STATEMENT = "Update CSMS_PO_hd set SUBLET_TOTAL_AMT = " & Sublet_TOTAL_AMT & ", SUBLET_TOTAL_VAT = " & Sublet_TOTAl_VAT & ", SUBLET_TOTAL_NET_AMT =" & Sublet_NET_AMT & " WHERE Po_No ='" & frmCSMS_PurchaseOrder.txtPoNumber.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "SUBLET PURCHASE", SQL_STATEMENT, FindTransactionID(vPONO, "PO_NO", "CSMS_PO_HD"), "", "PO NO: " & Null2String(vPONO) & " - ADD DETAILS", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            Call PO_GetID(frmCSMS_PurchaseOrder.txtPoNumber.Text)
            frmCSMS_PurchaseOrder.labDET.Caption = ""
            Unload Me
            'frmCSMS_PurchaseOrder.EnabledFrame (True)
        End If
    End If
End Sub
Function checkifvatornonvat(xxx As Variant) As String
     Dim rsnonvatwithvat                                 As New ADODB.Recordset
     Dim xtechcode                                       As String
        Set rsnonvatwithvat = New ADODB.Recordset
        Set rsnonvatwithvat = gconDMIS.Execute("Select nonvat from all_vendor_table where code = " & xxx & " ")
        If Not rsnonvatwithvat.EOF And Not rsnonvatwithvat.BOF Then
            checkifvatornonvat = rsnonvatwithvat!nonvat
        End If
    Set rsnonvatwithvat = Nothing
End Function

