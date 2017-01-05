VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSMIS_Inquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES MANAGEMENT INFORMATION SYSTEM'S QUERY"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15540
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Inquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15540
   Begin VB.PictureBox PICFILTER 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   15405
      TabIndex        =   21
      Top             =   8190
      Visible         =   0   'False
      Width           =   15435
      Begin VB.ComboBox cboSearchBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Inquiry.frx":030A
         Left            =   150
         List            =   "Inquiry.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
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
         Left            =   15000
         TabIndex        =   27
         Top             =   30
         Width           =   345
      End
      Begin VB.TextBox txtFilter_VI 
         Enabled         =   0   'False
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
         Left            =   5490
         TabIndex        =   23
         ToolTipText     =   "Add Filter Vehicle Invoice Number"
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox txtFilter_CS 
         Enabled         =   0   'False
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
         Left            =   2850
         TabIndex        =   22
         ToolTipText     =   "Add Filter Conduction Sticker Number"
         Top             =   270
         Width           =   1605
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CS#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2130
         TabIndex        =   24
         Top             =   270
         Width           =   2265
      End
      Begin VB.OptionButton Option2 
         Caption         =   "VI #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   25
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Search Method"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Click On Option To Add FIlter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2850
         TabIndex        =   26
         Top             =   30
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   795
      Left            =   13980
      MouseIcon       =   "Inquiry.frx":030E
      MousePointer    =   99  'Custom
      Picture         =   "Inquiry.frx":0460
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancel"
      Top             =   60
      Width           =   705
   End
   Begin VB.ComboBox cboYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "Inquiry.frx":079E
      Left            =   9510
      List            =   "Inquiry.frx":07A0
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   503
      Width           =   1185
   End
   Begin VB.PictureBox picTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   11880
      ScaleHeight     =   735
      ScaleWidth      =   1395
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   1395
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label labTot 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL RESULT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   0
      Width           =   11595
      Begin VB.OptionButton optSalesPer 
         Caption         =   "SAE Performance"
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
         Left            =   9300
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   2295
      End
      Begin VB.OptionButton optVehStock 
         Caption         =   "Vehicles On Stock"
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
         Left            =   7110
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   2205
      End
      Begin VB.OptionButton optCarRelease 
         Caption         =   "Total Car Release"
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
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   2145
      End
      Begin VB.OptionButton optInvCars 
         Caption         =   "Invoiced Cars"
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
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   90
         Width           =   1725
      End
      Begin VB.OptionButton optAllCars 
         Caption         =   "Allocated (Unit Recd) Vehicles"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   3165
      End
   End
   Begin VB.ComboBox cboModel 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   345
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   503
      Width           =   4095
   End
   Begin VB.ComboBox cboSalesAE 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   345
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   503
      Width           =   4125
   End
   Begin VB.TextBox txtMCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2EEE9&
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   10920
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   540
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdInquire 
      Caption         =   "&Inquiry"
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
      Left            =   13290
      MouseIcon       =   "Inquiry.frx":07A2
      MousePointer    =   99  'Custom
      Picture         =   "Inquiry.frx":08F4
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Search"
      Top             =   60
      Width           =   705
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   345
      Left            =   6630
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   503
      Width           =   2325
   End
   Begin MSFlexGridLib.MSFlexGrid grdInquiry 
      Height          =   8085
      Left            =   60
      TabIndex        =   20
      Top             =   900
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   14261
      _Version        =   393216
      Cols            =   8
      BackColor       =   16777215
      BackColorFixed  =   12648447
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      GridColor       =   8421504
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Inquiry.frx":0C3B
   End
   Begin VB.Label labYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8970
      TabIndex        =   16
      Top             =   578
      Width           =   645
   End
   Begin VB.Label labModel 
      BackStyle       =   0  'Transparent
      Caption         =   "By Model"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   570
      Width           =   840
   End
   Begin VB.Label labMonth 
      BackStyle       =   0  'Transparent
      Caption         =   "By Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   5820
      TabIndex        =   13
      Top             =   570
      Width           =   750
   End
   Begin VB.Label labSalesAE 
      BackStyle       =   0  'Transparent
      Caption         =   "By SAE Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   90
      TabIndex        =   7
      Top             =   570
      Width           =   1245
   End
End
Attribute VB_Name = "frmSMIS_Inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV2                                                         As ADODB.Recordset
Dim GCOL, GSORT
Attribute GSORT.VB_VarUserMemId = 1073938433

Sub FillGrid()
    Dim rsSO                                                          As ADODB.Recordset
    Dim RSPODAY                                                       As ADODB.Recordset
    Dim MANTH                                                         As Integer
    Dim cnt                                                           As Integer
    Dim PULLOUTAGE, AGEREC, AGEPO, DATEPO
    Dim SAE, TERM, INSCOMP, INSDATE, FINCOMP, INSUREDDATE, AGESOLD, BANKTERM, VDRNO, CustName, CustAdd
    Dim MODELX                                                        As String


    Dim i                                                             As Long
    On Error GoTo ErrorCode:
    cnt = 0
    If Not cboModel = "ALL" Then
        MODELX = "'" & cboModel & "'"
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''ALLOCATED CARS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optAllCars.Value = True Then
        LogAudit "V", "ALLOCATED CAR INQUIRY"
        Set rsMRRINV2 = New ADODB.Recordset
        If MODELX = "" Then
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE month(datereceived) = " & MANTH & " AND year(datereceived) = " & cboYear.Text & " AND STATUS='P' order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE month(datereceived) = " & MANTH & " AND STATUS='P' order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE year(datereceived) = " & cboYear.Text & " AND STATUS='P' order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

            Else
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE STATUS='P'  order by datereceived asc ", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE month(datereceived) = " & MANTH & " AND year(datereceived) = " & cboYear.Text & " AND STATUS='P' And Model = " & MODELX & " order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE month(datereceived) = " & MANTH & " AND STATUS='P' And Model = " & MODELX & "  order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE year(datereceived) = " & cboYear.Text & " AND STATUS='P' And Model = " & MODELX & "  order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsMRRINV2.Open "select ignkey as IGNKEY_NO , * from SMIS_MrrInv_table WHERE STATUS='P' And Model = " & MODELX & "  order by datereceived asc ", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If

        End If

        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                AGEPO = 0: DATEPO = "": AGEREC = 0: SAE = "": TERM = "": INSCOMP = "": INSDATE = "": FINCOMP = "": CustName = "": VDRNO = "": CustAdd = ""
                BANKTERM = ""

                If rsMRRINV2!ISTATUS = "R" Then

                    If IsDate(rsMRRINV2!PullOutDate) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), rsMRRINV2!DateReleased)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), rsMRRINV2!DateReleased)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    If IsDate(rsMRRINV2!datereceived) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), rsMRRINV2!DateReleased)
                    End If

                ElseIf rsMRRINV2!ISTATUS = "S" Then
                    If IsDate(rsMRRINV2!PullOutDate) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), rsMRRINV2!InvoicedDate)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), rsMRRINV2!InvoicedDate)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    If IsDate(rsMRRINV2!datereceived) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), rsMRRINV2!InvoicedDate)
                    End If
                Else
                    If IsDate(rsMRRINV2!PullOutDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), LOGDATE)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), LOGDATE)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If

                    If IsDate(rsMRRINV2!datereceived) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), LOGDATE)
                    End If
                End If



                If Null2String(rsMRRINV2!VI_NO) <> "" Then
                    Set rsSO = gconDMIS.Execute("SELECT HomeAddress, terms, VDR_NO, CustName, SALESAE,  FINANCINGCO , INSURANCECOMPANY  , INSUREDDATE  , TERM   FROM SMIS_SALESORDER WHERE IGNKEY_NO =" & N2Str2Null(rsMRRINV2!ignkey))
                    If Not rsSO.EOF Or Not rsSO.BOF Then
                        SAE = UCase(Null2String(rsSO!salesae))
                        TERM = UCase(Null2String(rsSO!TERM))
                        INSCOMP = UCase(Null2String(rsSO!INSURANCECOMPANY))
                        INSDATE = Format(Null2String(rsSO!INSUREDDATE), "MM/DD/YYYY")
                        FINCOMP = UCase(Null2String(rsSO!financingco))
                        CustName = UCase(Null2String(rsSO!CustName))
                        VDRNO = Null2String(rsSO!VDR_NO)
                        CustAdd = UCase(Null2String(rsSO!HomeAddress))
                        BANKTERM = N2Str2IntZero(rsSO!TERMS)
                    End If

                End If
                grdInquiry.AddItem UCase(rsMRRINV2!DESCRIPT) & Chr(9) & _
                                   Null2String(rsMRRINV2!ignkey) & Chr(9) & _
                                   Format(DATEPO, "mm/dd/yyyy") & Chr(9) & Null2String(rsMRRINV2!PONO) & Chr(9) & _
                                   Format(rsMRRINV2!datereceived, "mm/dd/yyyy") & Chr(9) & Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy") & Chr(9) & Null2String(rsMRRINV2!Source) & Chr(9) & _
                                   Null2String(rsMRRINV2!refPONO) & Chr(9) & Null2String(rsMRRINV2!drno) & Chr(9) & _
                                   rsMRRINV2!VI_NO & Chr(9) & VDRNO & Chr(9) & _
                                   Format(rsMRRINV2!InvoicedDate, "mm/dd/yyyy") & Chr(9) & Format(rsMRRINV2!DateReleased, "mm/dd/yyyy") & Chr(9) & _
                                   Null2String(rsMRRINV2!Color) & Chr(9) & _
                                   Null2String(rsMRRINV2!Vino) & Chr(9) & Null2String(rsMRRINV2!EngineNo) & Chr(9) & Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                   Null2String(rsMRRINV2!ISTATUS) & Chr(9) & _
                                   AGEPO & Chr(9) & AGEREC & Chr(9) & PULLOUTAGE & Chr(9) & _
                                   SAE & Chr(9) & _
                                   TERM & Chr(9) & FINCOMP & Chr(9) & BANKTERM & Chr(9) & _
                                   INSCOMP & Chr(9) & INSDATE & Chr(9) & _
                                   CustName & Chr(9) & CustAdd

                cnt = cnt + 1




                If cnt = 1 Then grdInquiry.RemoveItem 1
                rsMRRINV2.MoveNext
            Loop
            txtTotal.Text = cnt
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''INVOICED CARS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optInvCars.Value = True Then
        LogAudit "V", "INVOICED CAR INQUIRY"
        Set rsMRRINV2 = New ADODB.Recordset
        If MODELX = "" Then
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  month(invoiceddate) = " & MANTH & " AND year(invoiceddate) = " & cboYear.Text & " AND  ISNULL(INVOICEDDATE,0)<>0  order by invoiceddate,model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  year(invoiceddate) = " & cboYear.Text & " AND  ISNULL(INVOICEDDATE,0)<>0  order by invoiceddate,model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  month(invoiceddate) = " & MANTH & " AND ISNULL(INVOICEDDATE,0)<>0  order by invoiceddate,model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null  AND  ISNULL(INVOICEDDATE,0)<>0   order by invoiceddate , model  asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  month(invoiceddate) = " & MANTH & " AND year(invoiceddate) = " & cboYear.Text & " AND  ISNULL(INVOICEDDATE,0)<>0  AND MODEL=" & MODELX & " order by invoiceddate , model  asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  year(invoiceddate) = " & cboYear.Text & " AND  ISNULL(INVOICEDDATE,0)<>0  AND MODEL=" & MODELX & " order by invoiceddate , model  asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                MANTH = What_month(cboMonth.Text)
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null AND  month(invoiceddate) = " & MANTH & " AND ISNULL(INVOICEDDATE,0)<>0  AND MODEL=" & MODELX & " order by invoiceddate , model  asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE VI_NO is Not Null  AND  ISNULL(INVOICEDDATE,0)<>0   AND MODEL=" & MODELX & " order by invoiceddate , model  asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If

        End If



        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            INSDATE = "N/A"

            Do While Not rsMRRINV2.EOF
                If Null2String(rsMRRINV2!STATUS) <> "C" Or Null2String(rsMRRINV2!SOSTATUS) <> "C" Then
                    If IsDate((rsMRRINV2!INSUREDDATE)) = True Then
                        INSDATE = DateAdd("Y", 1, Null2Date(rsMRRINV2!INSUREDDATE))
                    End If

                    grdInquiry.AddItem Format(Null2String(rsMRRINV2!InvoicedDate), "MM/DD/YYYY") & Chr(9) & _
                                       Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                       Null2String(rsMRRINV2!VI_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                       Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!frameno) & Chr(9) & _
                                       Null2String(rsMRRINV2!CustName) & Chr(9) & _
                                       Null2String(rsMRRINV2!Color) & Chr(9) & _
                                       Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                       Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                       Null2String(rsMRRINV2!financingco) & Chr(9) & _
                                       Null2String(rsMRRINV2!TERMS) & Chr(9) & _
                                       IIf(Null2String(rsMRRINV2!INSURANCECOMPANY) = "", "OWN", rsMRRINV2!INSURANCECOMPANY) & Chr(9) & _
                                       INSDATE & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                       ""
                    cnt = cnt + 1
                    txtTotal.Text = cnt


                End If
                If cnt = 1 Then grdInquiry.RemoveItem 1

                rsMRRINV2.MoveNext

            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''RELEASED CARS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optCarRelease.Value = True Then
        LogAudit "V", "RELEASED CAR INQUIRY"
        Set rsMRRINV2 = New ADODB.Recordset
        If MODELX = "" Then
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND year(datereleased)= " & cboYear & " and month(datereleased)=" & What_month(cboMonth) & "  order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND year(datereleased)= " & cboYear & " order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND month(datereleased)=" & What_month(cboMonth) & "  order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE  ISDATE(datereleased)=1 order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If

        Else
            If cboMonth.Text <> "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND year(datereleased)= " & cboYear & " and month(datereleased)=" & What_month(cboMonth) & "  and model=" & MODELX & "  order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text = "ALL" And cboYear <> "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND year(datereleased)= " & cboYear & "  and model=" & MODELX & "    order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            ElseIf cboMonth.Text <> "ALL" And cboYear = "ALL" Then
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE ISDATE(datereleased)=1  AND month(datereleased)=" & What_month(cboMonth) & "  and model=" & MODELX & "  order by datereleased , model asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsMRRINV2.Open "select * from SMIS_PurchAgree WHERE  ISDATE(datereleased)=1 and model=" & MODELX & "  order by datereleased , model asc ", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If

        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                If IsDate((rsMRRINV2!INSUREDDATE)) = True Then
                    INSDATE = DateAdd("Y", 1, Null2Date(rsMRRINV2!INSUREDDATE))
                End If


                grdInquiry.AddItem Format(Null2String(rsMRRINV2!InvoicedDate), "MM/DD/YYYY") & Chr(9) & _
                                   Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                   Null2String(rsMRRINV2!VI_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                   Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!frameno) & Chr(9) & _
                                   Null2String(rsMRRINV2!CustName) & Chr(9) & _
                                   Null2String(rsMRRINV2!Color) & Chr(9) & _
                                   Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                   Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                   Null2String(rsMRRINV2!financingco) & Chr(9) & _
                                   Null2String(rsMRRINV2!TERMS) & Chr(9) & _
                                   IIf(Null2String(rsMRRINV2!INSURANCECOMPANY) = "", "OWN", rsMRRINV2!INSURANCECOMPANY) & Chr(9) & _
                                   INSDATE & Chr(9) & _
                                   Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                   Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                   ""


                cnt = cnt + 1
                txtTotal.Text = cnt
                rsMRRINV2.MoveNext
                If cnt = 1 Then grdInquiry.RemoveItem 1
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ON STOCK
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If optVehStock.Value = True Then
        LogAudit "V", "VEHICLE ON STOCK INQUIRY"
        If cboModel.Text <> "ALL" Then
            Set rsMRRINV2 = New ADODB.Recordset
            rsMRRINV2.Open "select ignkey as IGNKEY_NO ,  * from SMIS_MRRINV_TABLE WHERE Model = '" & cboModel.Text & "' and released=0  and status='P'  order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            
    ' ***** REVISED BY: DHANG_ERZ
    
        Set rsMRRINV2 = New ADODB.Recordset
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "JMC" Then
            rsMRRINV2.Open "select ignkey as IGNKEY_NO ,  * from SMIS_MrrInv_Table A where released=0  and A.Status='P' AND A.ignkey NOT IN (SELECT IGNKEY_NO FROM SMIS_SalesOrder WHERE STATUS<>'C') order by datereceived asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsMRRINV2.Open "SELECT IGNKEY AS IGNKEY_NO ,  * FROM SMIS_MRRINV_TABLE WHERE RELEASED=0  AND STATUS='P' ORDER BY DATERECEIVED ASC  ", gconDMIS, adOpenForwardOnly, adLockReadOnly
                      
        End If
  
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                If rsMRRINV2!RELEASED = False Then
                    If IsDate(rsMRRINV2!PullOutDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), LOGDATE)
                    Else
                        PULLOUTAGE = 0
                    End If

                    AGEPO = 0
                    DATEPO = ""

                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), LOGDATE)
                                DATEPO = RSPODAY!DATEORDERED
                            End If
                        End If
                    End If
                    AGEREC = 0
                    If IsDate(rsMRRINV2!datereceived) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), LOGDATE)
                    Else
                        AGEREC = 0
                    End If


                    grdInquiry.AddItem Null2String(rsMRRINV2!DESCRIPT) & Chr(9) & _
                                       Null2String(rsMRRINV2!ignkey) & Chr(9) & _
                                       Format(DATEPO, "mm/dd/yyyy") & Chr(9) & _
                                       Null2String(rsMRRINV2!PONO) & Chr(9) & _
                                       Null2String(Format(rsMRRINV2!datereceived, "MM/DD/YYYY")) & Chr(9) & _
                                       Null2String(Format(rsMRRINV2!PullOutDate, "MM/DD/YYYY")) & Chr(9) & _
                                       Null2String(rsMRRINV2!Source) & Chr(9) & _
                                       Null2String(rsMRRINV2!refPONO) & Chr(9) & _
                                       Null2String(rsMRRINV2!drno) & Chr(9) & _
                                       Null2String(rsMRRINV2!Color) & Chr(9) & _
                                       Null2String(rsMRRINV2!Vino) & Chr(9) & _
                                       Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       AGEPO & Chr(9) & AGEREC & Chr(9) & PULLOUTAGE & Chr(9) & _
                                       Null2String(rsMRRINV2!LTOStatus) & Chr(9) & _
                                       Null2String(rsMRRINV2!CSR) & Chr(9) & _
                                       Null2String(rsMRRINV2!CSRDATE) & Chr(9) & _
                                       Null2String(rsMRRINV2!Location)
                    cnt = cnt + 1
                    txtTotal.Text = cnt
                    'DoEvents
                    If cnt = 1 Then grdInquiry.RemoveItem 1
                End If
                rsMRRINV2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'SAE PERFORMANCE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If optSalesPer.Value = True Then
        LogAudit "V", "SAE PERFORMANCE"
        If cboSalesAE.Text = "ALL" Then
            Set rsMRRINV2 = New ADODB.Recordset
            rsMRRINV2.Open "select * from SMIS_SALESORDER WHERE   status='P'  order by salesae asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            Set rsMRRINV2 = New ADODB.Recordset
            rsMRRINV2.Open "select * from SMIS_SALESORDER WHERE salesae = '" & cboSalesAE.Text & "' and status='P' order by datereleased asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                AGESOLD = 0: INSUREDDATE = ""

                If Null2String(rsMRRINV2!InvoicedDate) <> "" Then
                    AGESOLD = DateDiff("d", Null2String(rsMRRINV2!InvoicedDate), Null2String(rsMRRINV2!DEYT))
                End If

                If Null2String(rsMRRINV2!INSURANCECOMPANY) <> "" Then
                    INSUREDDATE = Null2String(rsMRRINV2!INSUREDDATE)
                End If
                If Null2String(rsMRRINV2!DateReleased) <> "" Then

                    grdInquiry.AddItem Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                       Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                       Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                       Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       Null2String(rsMRRINV2!Vino) & Chr(9) & AGESOLD & Chr(9) & _
                                       Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                       Null2String(rsMRRINV2!VI_NO) & Chr(9) & Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!financingco) & Chr(9) & Null2String(rsMRRINV2!INSURANCECOMPANY) & Chr(9) & INSUREDDATE & Chr(9) & Null2String(rsMRRINV2!Color)

                    cnt = cnt + 1
                    txtTotal.Text = cnt

                    If cnt = 1 Then grdInquiry.RemoveItem 1
                End If
                rsMRRINV2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If





    Exit Sub
ErrorCode:
    ShowVBError
   End If
End Sub

Sub initGrid()
    labSalesAE.Visible = False
    cboSalesAE.Visible = False
    labMonth.Visible = False
    cboMonth.Visible = False
    labYear.Visible = False
    cboYear.Visible = False
    labModel.Visible = False
    cboModel.Visible = False

    txtTotal.Text = "0"
    grdInquiry.Height = 8085
    ShowHidePictureBox2 PICFILTER, True

    If optAllCars.Value = True Then
        With grdInquiry
            .FixedCols = 2: .Cols = 29: .Row = 0
            .ColWidth(0) = 3000: .Col = 0: .Text = "Unit Description": .ColAlignment(0) = 1
            .ColWidth(1) = 800: .Col = 1: .Text = "CSNO": .ColAlignment(1) = 1
            .ColWidth(2) = 1000: .Col = 2: .Text = "PO Date": .ColAlignment(2) = 0
            .ColWidth(3) = 900: .Col = 3: .Text = "PO No": .ColAlignment(3) = 3
            .ColWidth(4) = 1000: .Col = 4: .Text = "Recieved": .ColAlignment(4) = 0
            .ColWidth(5) = 1000: .Col = 5: .Text = "PullOutDate": .ColAlignment(5) = 0
            .ColWidth(6) = 900: .Col = 6: .Text = "Source": .ColAlignment(6) = 0
            .ColWidth(7) = 800: .Col = 7: .Text = "RefInv#": .ColAlignment(7) = 3
            .ColWidth(8) = 800: .Col = 8: .Text = "RefDR#": .ColAlignment(8) = 3

            .ColWidth(9) = 800: .Col = 9: .Text = "VI#": .ColAlignment(9) = 3
            .ColWidth(10) = 800: .Col = 10: .Text = "VDR#": .ColAlignment(10) = 3
            .ColWidth(11) = 1000: .Col = 11: .Text = "Date Inv": .ColAlignment(11) = 0
            .ColWidth(12) = 1000: .Col = 12: .Text = "Date Rel": .ColAlignment(12) = 0
            .ColWidth(13) = 1700: .Col = 13: .Text = "Color": .ColAlignment(13) = 0
            .ColWidth(14) = 1800: .Col = 14: .Text = "VIN#": .ColAlignment(14) = 0
            .ColWidth(15) = 1400: .Col = 15: .Text = "Engine#": .ColAlignment(15) = 0
            .ColWidth(16) = 1400: .Col = 16: .Text = "Prod#": .ColAlignment(16) = 0

            .ColWidth(17) = 600: .Col = 17: .Text = "Status": .ColAlignment(17) = 3
            .ColWidth(18) = 1000: .Col = 18: .Text = "Age(PO)": .ColAlignment(18) = 3
            .ColWidth(19) = 1000: .Col = 19: .Text = "Age(Recd)": .ColAlignment(19) = 3
            .ColWidth(20) = 1000: .Col = 20: .Text = "Age(PullOut)": .ColAlignment(20) = 3

            .ColWidth(21) = 2100: .Col = 21: .Text = "Sales Agent": .ColAlignment(21) = 0
            .ColWidth(22) = 800: .Col = 22: .Text = "Term": .ColAlignment(22) = 3
            .ColWidth(23) = 1800: .Col = 23: .Text = "Financing Company": .ColAlignment(23) = 0
            .ColWidth(24) = 900: .Col = 24: .Text = "Bank Term": .ColAlignment(24) = 3
            .ColWidth(25) = 1800: .Col = 25: .Text = "Insurance Company": .ColAlignment(25) = 0
            .ColWidth(26) = 1000: .Col = 26: .Text = "Insured Date": .ColAlignment(26) = 0
            .ColWidth(27) = 3500: .Col = 27: .Text = "Customer Name": .ColAlignment(27) = 0
            .ColWidth(28) = 4500: .Col = 28: .Text = "Customer Address": .ColAlignment(28) = 0
        End With
        labModel.Visible = True: cboModel.Visible = True
        labMonth.Visible = True: cboMonth.Visible = True
        labYear.Visible = True: cboYear.Visible = True
        fillcbomunth
        FillCboMoreYear cboYear
        Call cboYear.AddItem("ALL", 0)
        On Error Resume Next
        cboYear.Text = Year(LOGDATE)
    End If

    If optInvCars.Value = True Then
        With grdInquiry
            .FixedCols = 5: .Row = 0: .Cols = 20
            .ColWidth(0) = 1000: .Col = 0: .Text = "Date Inv": .ColAlignment(0) = 1
            .ColWidth(1) = 1000: .Col = 1: .Text = "Date Rel": .ColAlignment(1) = 1
            .ColWidth(2) = 700: .Col = 2: .Text = "VI#": .ColAlignment(2) = 3
            .ColWidth(3) = 700: .Col = 3: .Text = "VDR#": .ColAlignment(3) = 3
            .ColWidth(4) = 3500: .Col = 4: .Text = "Unit Description": .ColAlignment(4) = 0
            .ColWidth(5) = 800: .Col = 5: .Text = "CS#": .ColAlignment(5) = 0
            .ColWidth(6) = 1800: .Col = 6: .Text = "VIN#": .ColAlignment(6) = 0

            .ColWidth(7) = 3000: .Col = 7: .Text = "Customer Name": .ColAlignment(7) = 0
            .ColWidth(8) = 1500: .Col = 8: .Text = "Color": .ColAlignment(8) = 0
            .ColWidth(9) = 2500: .Col = 9: .Text = "Sales Agent": .ColAlignment(9) = 0

            .ColWidth(10) = 800: .Col = 10: .Text = "Term": .ColAlignment(10) = 0
            .ColWidth(11) = 2000: .Col = 11: .Text = "Financing Company": .ColAlignment(11) = 0
            .ColWidth(12) = 1000: .Col = 12: .Text = "Bank Term": .ColAlignment(12) = 3
            .ColWidth(13) = 2000: .Col = 13: .Text = "Insurance Company": .ColAlignment(13) = 0
            .ColWidth(14) = 1000: .Col = 14: .Text = "Expires On": .ColAlignment(14) = 0
            .ColWidth(15) = 1500: .Col = 15: .Text = "Prod#": .ColAlignment(15) = 0
            .ColWidth(16) = 1500: .Col = 16: .Text = "Engine#": .ColAlignment(16) = 0
            .ColWidth(17) = 1500: .Col = 17: .Text = "CSR": .ColAlignment(17) = 0
            .ColWidth(18) = 1500: .Col = 18: .Text = "CSR DATE": .ColAlignment(18) = 0
            .ColWidth(19) = 1500: .Col = 19: .Text = "LTO Status": .ColAlignment(19) = 0

        End With
        labModel.Visible = True: cboModel.Visible = True
        labMonth.Visible = True: cboMonth.Visible = True
        labYear.Visible = True: cboYear.Visible = True
        fillcbomunth
        FillCboMoreYear cboYear
        Call cboYear.AddItem("ALL", 0)
        On Error Resume Next
        cboYear.Text = Year(LOGDATE)
    End If

    If optCarRelease.Value = True Then
        With grdInquiry
            .FixedCols = 5: .Row = 0: .Cols = 20
            .ColWidth(0) = 1000: .Col = 0: .Text = "Date Inv": .ColAlignment(0) = 1
            .ColWidth(1) = 1000: .Col = 1: .Text = "Date Rel": .ColAlignment(1) = 1
            .ColWidth(2) = 700: .Col = 2: .Text = "VI#": .ColAlignment(2) = 3
            .ColWidth(3) = 700: .Col = 3: .Text = "VDR#": .ColAlignment(3) = 3
            .ColWidth(4) = 3500: .Col = 4: .Text = "Unit Description": .ColAlignment(4) = 0
            .ColWidth(5) = 800: .Col = 5: .Text = "CS#": .ColAlignment(5) = 0
            .ColWidth(6) = 1800: .Col = 6: .Text = "VIN#": .ColAlignment(6) = 0

            .ColWidth(7) = 3000: .Col = 7: .Text = "Customer Name": .ColAlignment(7) = 0
            .ColWidth(8) = 1500: .Col = 8: .Text = "Color": .ColAlignment(8) = 0
            .ColWidth(9) = 2500: .Col = 9: .Text = "Sales Agent": .ColAlignment(9) = 0

            .ColWidth(10) = 800: .Col = 10: .Text = "Term": .ColAlignment(10) = 0
            .ColWidth(11) = 2000: .Col = 11: .Text = "Financing Company": .ColAlignment(11) = 0
            .ColWidth(12) = 1000: .Col = 12: .Text = "Bank Term": .ColAlignment(12) = 3
            .ColWidth(13) = 2000: .Col = 13: .Text = "Insurance Company": .ColAlignment(13) = 0
            .ColWidth(14) = 1000: .Col = 14: .Text = "Expires On": .ColAlignment(14) = 0
            .ColWidth(15) = 1500: .Col = 15: .Text = "Prod#": .ColAlignment(15) = 0
            .ColWidth(16) = 1500: .Col = 16: .Text = "Engine#": .ColAlignment(16) = 0
            .ColWidth(17) = 1500: .Col = 17: .Text = "CSR": .ColAlignment(17) = 0
            .ColWidth(18) = 1500: .Col = 18: .Text = "CSR DATE": .ColAlignment(18) = 0
            .ColWidth(19) = 1500: .Col = 19: .Text = "LTO Status": .ColAlignment(19) = 0
        End With
        labModel.Visible = True: cboModel.Visible = True
        labMonth.Visible = True: cboMonth.Visible = True
        labYear.Visible = True: cboYear.Visible = True
        fillcbomunth
        FillCboMoreYear cboYear
        Call cboYear.AddItem("ALL", 0)
        On Error Resume Next
        cboYear.Text = Year(LOGDATE)
    End If

    If optVehStock.Value = True Then

        With grdInquiry
            .FixedCols = 2: .Cols = 20: .Row = 0
            .ColWidth(0) = 3000: .Col = 0: .Text = "Unit Description": .ColAlignment(0) = 1
            .ColWidth(1) = 800: .Col = 1: .Text = "CSNO": .ColAlignment(1) = 1
            .ColWidth(2) = 1000: .Col = 2: .Text = "PO Date": .ColAlignment(2) = 0
            .ColWidth(3) = 900: .Col = 3: .Text = "PO No": .ColAlignment(3) = 3
            .ColWidth(4) = 1000: .Col = 4: .Text = "Recieved": .ColAlignment(4) = 0
            .ColWidth(5) = 1000: .Col = 5: .Text = "PullOutDate": .ColAlignment(5) = 0
            .ColWidth(6) = 900: .Col = 6: .Text = "Source": .ColAlignment(6) = 0
            .ColWidth(7) = 800: .Col = 7: .Text = "RefInv#": .ColAlignment(7) = 3
            .ColWidth(8) = 800: .Col = 8: .Text = "RefDR#": .ColAlignment(8) = 3


            .ColWidth(9) = 1700: .Col = 9: .Text = "Color": .ColAlignment(9) = 0
            .ColWidth(10) = 1800: .Col = 10: .Text = "VIN#": .ColAlignment(10) = 0
            .ColWidth(11) = 1400: .Col = 11: .Text = "Engine#": .ColAlignment(11) = 0
            .ColWidth(12) = 1400: .Col = 12: .Text = "Prod#": .ColAlignment(12) = 0


            .ColWidth(13) = 1000: .Col = 13: .Text = "Age(PO)": .ColAlignment(13) = 3
            .ColWidth(14) = 1000: .Col = 14: .Text = "Age(Recd)": .ColAlignment(14) = 3
            .ColWidth(15) = 1000: .Col = 15: .Text = "Age(PullOut)": .ColAlignment(15) = 3

            .ColWidth(16) = 2100: .Col = 16: .Text = "LTO Status": .ColAlignment(16) = 0
            .ColWidth(17) = 800: .Col = 17: .Text = "CSR": .ColAlignment(17) = 3
            .ColWidth(18) = 1800: .Col = 18: .Text = "CSR Date": .ColAlignment(18) = 0
            .ColWidth(19) = 1800: .Col = 19: .Text = "Location": .ColAlignment(19) = 0

        End With
        labModel.Visible = True
        cboModel.Visible = True

    End If


    If optSalesPer.Value = True Then
        With grdInquiry
            .FixedCols = 3: .Row = 0: .Cols = 14
            .ColWidth(0) = 1800: .Col = 0: .Text = "Sale Agent Name": .ColAlignment(0) = 1
            .ColWidth(1) = 3500: .Col = 1: .Text = "Model .": .ColAlignment(1) = 1
            .ColWidth(2) = 800: .Col = 2: .Text = "CSNo": .ColAlignment(2) = 0
            .ColWidth(3) = 800: .Col = 3: .Text = "Date Released": .ColAlignment(3) = 0
            .ColWidth(4) = 1500: .Col = 4: .Text = "Prod No.": .ColAlignment(4) = 0
            .ColWidth(5) = 1500: .Col = 5: .Text = "VIN No.": .ColAlignment(5) = 0
            .ColWidth(6) = 1500: .Col = 6: .Text = "Age Sold.": .ColAlignment(6) = 3
            .ColWidth(7) = 1000: .Col = 7: .Text = "Term.": .ColAlignment(7) = 3
            .ColWidth(8) = 1000: .Col = 8: .Text = "VI#.": .ColAlignment(8) = 0
            .ColWidth(9) = 1000: .Col = 9: .Text = "VDR#.": .ColAlignment(9) = 0
            .ColWidth(10) = 1000: .Col = 10: .Text = "Financing Company.": .ColAlignment(10) = 0
            .ColWidth(11) = 1100: .Col = 11: .Text = "Insurance Company.": .ColAlignment(11) = 0
            .ColWidth(12) = 1000: .Col = 12: .Text = "Insured Date.": .ColAlignment(12) = 0
            .ColWidth(13) = 1000: .Col = 13: .Text = "Color.": .ColAlignment(13) = 0


        End With
        labSalesAE.Visible = True
        cboSalesAE.Visible = True
        FillcboSAE
    End If

End Sub

Sub FillcboSAE()
    Dim rsSrep                                                        As ADODB.Recordset
    Set rsSrep = New ADODB.Recordset
    rsSrep.Open "select distinct salesae from smis_salesorder", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSrep.EOF And Not rsSrep.BOF Then
        cboSalesAE.Clear
        cboSalesAE.AddItem "ALL"
        cboSalesAE.Text = "ALL"
        rsSrep.MoveFirst
        Do While Not rsSrep.EOF
            cboSalesAE.AddItem Null2String(rsSrep!salesae)
            rsSrep.MoveNext
        Loop
    End If
End Sub

Sub fillcbomunth()
    cboMonth.Clear
    cboMonth.AddItem "ALL"
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    cboMonth.Text = "ALL"
End Sub

Sub FillCboModel()
    Dim rsMRRVEH                                                      As ADODB.Recordset
    Set rsMRRVEH = New ADODB.Recordset
    rsMRRVEH.Open "select DISTINCT upper(MODEL) as MODEL from ALL_MODEL WHERE ISNULL(MODEL,'')<>''", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRVEH.EOF And Not rsMRRVEH.BOF Then
        cboModel.Clear
        cboModel.AddItem "ALL", 0
        cboModel.Text = "ALL"
        rsMRRVEH.MoveFirst
        Do While Not rsMRRVEH.EOF
            cboModel.AddItem Null2String(rsMRRVEH!Model)
            rsMRRVEH.MoveNext
        Loop
    End If
End Sub

Sub FillCboModel2()
    Dim rsMRRVEH                                                      As ADODB.Recordset
    Set rsMRRVEH = New ADODB.Recordset
    rsMRRVEH.Open "select * from SMIS_MrrInv WHERE released = 0 order by descript asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRVEH.EOF And Not rsMRRVEH.BOF Then
        cboModel.Clear
        cboModel.AddItem "ALL"
        cboModel.Text = "ALL"
        rsMRRVEH.MoveFirst
        Do While Not rsMRRVEH.EOF
            cboModel.AddItem Null2String(rsMRRVEH!DESCRIPT)
            txtMCode.Text = Null2String(rsMRRVEH!Code)
            rsMRRVEH.MoveNext
        Loop
    End If
End Sub

Sub FILTERVIEW()
    Dim rsSO                                                          As ADODB.Recordset
    Dim RSPODAY                                                       As ADODB.Recordset
    Dim MANTH                                                         As Integer
    Dim cnt                                                           As Integer
    Dim PULLOUTAGE, AGEREC, AGEPO, DATEPO
    Dim SAE, TERM, INSCOMP, INSDATE, FINCOMP, INSUREDDATE, AGESOLD, BANKTERM, VDRNO, CustName, CustAdd
    Dim MODELX                                                        As String


    Dim i                                                             As Long
    On Error GoTo ErrorCode:
    cnt = 0

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''ALLOCATED CARS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optAllCars.Value = True Then
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                AGEPO = 0: DATEPO = "": AGEREC = 0: SAE = "": TERM = "": INSCOMP = "": INSDATE = "": FINCOMP = "": CustName = "": VDRNO = "": CustAdd = ""
                BANKTERM = ""

                If rsMRRINV2!ISTATUS = "R" Then

                    If IsDate(rsMRRINV2!PullOutDate) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), rsMRRINV2!DateReleased)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), rsMRRINV2!DateReleased)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    If IsDate(rsMRRINV2!datereceived) = True And IsDate(rsMRRINV2!DateReleased) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), rsMRRINV2!DateReleased)
                    End If

                ElseIf rsMRRINV2!ISTATUS = "S" Then
                    If IsDate(rsMRRINV2!PullOutDate) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), rsMRRINV2!InvoicedDate)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), rsMRRINV2!InvoicedDate)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    If IsDate(rsMRRINV2!datereceived) = True And IsDate(rsMRRINV2!InvoicedDate) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), rsMRRINV2!InvoicedDate)
                    End If
                Else
                    If IsDate(rsMRRINV2!PullOutDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), LOGDATE)
                    End If
                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), LOGDATE)
                                DATEPO = Format(RSPODAY!DATEORDERED, "mm/dd/yyyy")
                            End If
                        End If
                    End If

                    If IsDate(rsMRRINV2!datereceived) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), LOGDATE)
                    End If
                End If



                If Null2String(rsMRRINV2!VI_NO) <> "" Then
                    Set rsSO = gconDMIS.Execute("SELECT HomeAddress, terms, VDR_NO, CustName, SALESAE,  FINANCINGCO , INSURANCECOMPANY  , INSUREDDATE  , TERM   FROM SMIS_SALESORDER WHERE STATUS='P' AND IGNKEY_NO =" & N2Str2Null(rsMRRINV2!ignkey))
                    If Not rsSO.EOF Or Not rsSO.BOF Then
                        SAE = UCase(Null2String(rsSO!salesae))
                        TERM = UCase(Null2String(rsSO!TERM))
                        INSCOMP = UCase(Null2String(rsSO!INSURANCECOMPANY))
                        INSDATE = Format(Null2String(rsSO!INSUREDDATE), "MM/DD/YYYY")
                        FINCOMP = UCase(Null2String(rsSO!financingco))
                        CustName = UCase(Null2String(rsSO!CustName))
                        VDRNO = Null2String(rsSO!VDR_NO)
                        CustAdd = UCase(Null2String(rsSO!HomeAddress))
                        BANKTERM = N2Str2IntZero(rsSO!TERMS)
                    End If

                End If
                grdInquiry.AddItem UCase(rsMRRINV2!DESCRIPT) & Chr(9) & _
                                   Null2String(rsMRRINV2!ignkey) & Chr(9) & _
                                   Format(DATEPO, "mm/dd/yyyy") & Chr(9) & Null2String(rsMRRINV2!PONO) & Chr(9) & _
                                   Format(rsMRRINV2!datereceived, "mm/dd/yyyy") & Chr(9) & Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy") & Chr(9) & Null2String(rsMRRINV2!Source) & Chr(9) & _
                                   Null2String(rsMRRINV2!refPONO) & Chr(9) & Null2String(rsMRRINV2!drno) & Chr(9) & _
                                   rsMRRINV2!VI_NO & Chr(9) & VDRNO & Chr(9) & _
                                   Format(rsMRRINV2!DateReleased, "mm/dd/yyyy") & Chr(9) & Format(rsMRRINV2!InvoicedDate, "mm/dd/yyyy") & Chr(9) & _
                                   Null2String(rsMRRINV2!Color) & Chr(9) & _
                                   Null2String(rsMRRINV2!Vino) & Chr(9) & Null2String(rsMRRINV2!EngineNo) & Chr(9) & Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                   Null2String(rsMRRINV2!ISTATUS) & Chr(9) & _
                                   AGEPO & Chr(9) & AGEREC & Chr(9) & PULLOUTAGE & Chr(9) & _
                                   SAE & Chr(9) & _
                                   TERM & Chr(9) & FINCOMP & Chr(9) & BANKTERM & Chr(9) & _
                                   INSCOMP & Chr(9) & INSDATE & Chr(9) & _
                                   CustName & Chr(9) & CustAdd

                cnt = cnt + 1
                txtTotal.Text = cnt



                If cnt = 1 Then grdInquiry.RemoveItem 1
                rsMRRINV2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''INVOICED CARS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optInvCars.Value = True Then
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            INSDATE = "N/A"

            Do While Not rsMRRINV2.EOF
                If Null2String(rsMRRINV2!STATUS) <> "C" Or Null2String(rsMRRINV2!SOSTATUS) <> "C" Then
                    If IsDate((rsMRRINV2!INSUREDDATE)) = True Then
                        INSDATE = DateAdd("Y", 1, Null2Date(rsMRRINV2!INSUREDDATE))
                    End If

                    grdInquiry.AddItem Format(Null2String(rsMRRINV2!InvoicedDate), "MM/DD/YYYY") & Chr(9) & _
                                       Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                       Null2String(rsMRRINV2!VI_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                       Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!frameno) & Chr(9) & _
                                       Null2String(rsMRRINV2!CustName) & Chr(9) & _
                                       Null2String(rsMRRINV2!Color) & Chr(9) & _
                                       Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                       Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                       Null2String(rsMRRINV2!financingco) & Chr(9) & _
                                       Null2String(rsMRRINV2!TERMS) & Chr(9) & _
                                       IIf(Null2String(rsMRRINV2!INSURANCECOMPANY) = "", "OWN", rsMRRINV2!INSURANCECOMPANY) & Chr(9) & _
                                       INSDATE & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                       ""
                    cnt = cnt + 1
                    txtTotal.Text = cnt


                End If
                If cnt = 1 Then grdInquiry.RemoveItem 1

                rsMRRINV2.MoveNext

            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''RELEASED CARS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optCarRelease.Value = True Then
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                If IsDate((rsMRRINV2!INSUREDDATE)) = True Then
                    INSDATE = DateAdd("Y", 1, Null2Date(rsMRRINV2!INSUREDDATE))
                End If


                grdInquiry.AddItem Format(Null2String(rsMRRINV2!InvoicedDate), "MM/DD/YYYY") & Chr(9) & _
                                   Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                   Null2String(rsMRRINV2!VI_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                   Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                   Null2String(rsMRRINV2!frameno) & Chr(9) & _
                                   Null2String(rsMRRINV2!CustName) & Chr(9) & _
                                   Null2String(rsMRRINV2!Color) & Chr(9) & _
                                   Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                   Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                   Null2String(rsMRRINV2!financingco) & Chr(9) & _
                                   Null2String(rsMRRINV2!TERMS) & Chr(9) & _
                                   IIf(Null2String(rsMRRINV2!INSURANCECOMPANY) = "", "OWN", rsMRRINV2!INSURANCECOMPANY) & Chr(9) & _
                                   INSDATE & Chr(9) & _
                                   Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                   Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                   ""


                cnt = cnt + 1
                txtTotal.Text = cnt
                rsMRRINV2.MoveNext
                If cnt = 1 Then grdInquiry.RemoveItem 1
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ON STOCK
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If optVehStock.Value = True Then
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                If rsMRRINV2!RELEASED = False Then
                    If IsDate(rsMRRINV2!PullOutDate) = True Then
                        PULLOUTAGE = DateDiff("d", Null2String(rsMRRINV2!PullOutDate), LOGDATE)
                    Else
                        PULLOUTAGE = 0
                    End If

                    AGEPO = 0
                    DATEPO = ""

                    If Null2String(rsMRRINV2!PONO) <> "" Then
                        Set RSPODAY = gconDMIS.Execute("SELECT  DATEORDERED  FROM SMIS_PO WHERE PO_NO='" & rsMRRINV2!PONO & "'")
                        If Not RSPODAY.EOF Or RSPODAY.BOF Then
                            If IsDate(RSPODAY!DATEORDERED) = True Then
                                AGEPO = DateDiff("d", Null2String(RSPODAY!DATEORDERED), LOGDATE)
                                DATEPO = RSPODAY!DATEORDERED
                            End If
                        End If
                    End If
                    AGEREC = 0
                    If IsDate(rsMRRINV2!datereceived) = True Then
                        AGEREC = DateDiff("d", Null2String(rsMRRINV2!datereceived), LOGDATE)
                    Else
                        AGEREC = 0
                    End If


                    grdInquiry.AddItem Null2String(rsMRRINV2!DESCRIPT) & Chr(9) & _
                                       Null2String(rsMRRINV2!ignkey) & Chr(9) & _
                                       Format(DATEPO, "mm/dd/yyyy") & Chr(9) & _
                                       Null2String(rsMRRINV2!PONO) & Chr(9) & _
                                       Null2String(Format(rsMRRINV2!datereceived, "MM/DD/YYYY")) & Chr(9) & _
                                       Null2String(Format(rsMRRINV2!PullOutDate, "MM/DD/YYYY")) & Chr(9) & _
                                       Null2String(rsMRRINV2!Source) & Chr(9) & _
                                       Null2String(rsMRRINV2!refPONO) & Chr(9) & _
                                       Null2String(rsMRRINV2!drno) & Chr(9) & _
                                       Null2String(rsMRRINV2!Color) & Chr(9) & _
                                       Null2String(rsMRRINV2!Vino) & Chr(9) & _
                                       Null2String(rsMRRINV2!EngineNo) & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       AGEPO & Chr(9) & AGEREC & Chr(9) & PULLOUTAGE & Chr(9) & _
                                       Null2String(rsMRRINV2!LTOStatus) & Chr(9) & _
                                       Null2String(rsMRRINV2!CSR) & Chr(9) & _
                                       Null2String(rsMRRINV2!CSRDATE)
                    cnt = cnt + 1
                    txtTotal.Text = cnt
                    'DoEvents
                    If cnt = 1 Then grdInquiry.RemoveItem 1
                End If
                rsMRRINV2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'SAE PERFORMANCE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If optSalesPer.Value = True Then
        If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
            Screen.MousePointer = 11
            rsMRRINV2.MoveFirst
            Do While Not rsMRRINV2.EOF
                AGESOLD = 0: INSUREDDATE = ""

                If Null2String(rsMRRINV2!InvoicedDate) <> "" Then
                    AGESOLD = DateDiff("d", Null2String(rsMRRINV2!InvoicedDate), Null2String(rsMRRINV2!DEYT))
                End If

                If Null2String(rsMRRINV2!INSURANCECOMPANY) <> "" Then
                    INSUREDDATE = Null2String(rsMRRINV2!INSUREDDATE)
                End If
                If Null2String(rsMRRINV2!DateReleased) <> "" Then

                    grdInquiry.AddItem Null2String(rsMRRINV2!salesae) & Chr(9) & _
                                       Null2String(rsMRRINV2!modeldescription) & Chr(9) & _
                                       Null2String(rsMRRINV2!IGNKEY_NO) & Chr(9) & _
                                       Format(Null2String(rsMRRINV2!DateReleased), "MM/DD/YYYY") & Chr(9) & _
                                       Null2String(rsMRRINV2!prodno) & Chr(9) & _
                                       Null2String(rsMRRINV2!Vino) & Chr(9) & AGESOLD & Chr(9) & _
                                       Null2String(rsMRRINV2!TERM) & Chr(9) & _
                                       Null2String(rsMRRINV2!VI_NO) & Chr(9) & Null2String(rsMRRINV2!VDR_NO) & Chr(9) & _
                                       Null2String(rsMRRINV2!financingco) & Chr(9) & Null2String(rsMRRINV2!INSURANCECOMPANY) & Chr(9) & INSUREDDATE & Chr(9) & Null2String(rsMRRINV2!Color)

                    cnt = cnt + 1
                    txtTotal.Text = cnt

                    If cnt = 1 Then grdInquiry.RemoveItem 1
                End If
                rsMRRINV2.MoveNext
            Loop
            Screen.MousePointer = 0
        End If
        Exit Sub
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub SHOWFILTERVIEW()

    txtTotal = 0
    cleargrid grdInquiry
    If txtFilter_CS.Enabled = True Then
        If txtFilter_CS = "" Then: Exit Sub
        If cboSearchBy.ListIndex = 0 Then
            rsMRRINV2.FILTER = "IGNKEY_NO LIKE '" & Repleys(txtFilter_CS.Text) & "%'"
        ElseIf cboSearchBy.ListIndex = 1 Then
            rsMRRINV2.FILTER = "IGNKEY_NO LIKE '%" & Repleys(txtFilter_CS.Text) & "%'"
        ElseIf cboSearchBy.ListIndex = 2 Then
            rsMRRINV2.FILTER = "IGNKEY_NO ='" & Repleys(txtFilter_CS) & "'"
        Else
            rsMRRINV2.FILTER = "IGNKEY_NO LIKE '" & Repleys(txtFilter_CS) & "%'"
        End If

    ElseIf txtFilter_VI.Enabled = True Then
        If txtFilter_VI = "" Then: Exit Sub
        If cboSearchBy.ListIndex = 0 Then
            rsMRRINV2.FILTER = "VI_NO like '" & Repleys(txtFilter_VI) & "%'"
        ElseIf cboSearchBy.ListIndex = 1 Then
            rsMRRINV2.FILTER = "VI_NO like '%" & Repleys(txtFilter_VI) & "%'"
        ElseIf cboSearchBy.ListIndex = 2 Then
            rsMRRINV2.FILTER = "VI_NO = '" & Repleys(txtFilter_VI) & "'"
        Else
            rsMRRINV2.FILTER = "VI_NO like '" & Repleys(txtFilter_VI) & "%'"
        End If

    End If
    FILTERVIEW
End Sub

Private Sub cmdInquire_Click()
    Form_KeyDown vbKeyEscape, 0
    txtTotal = 0
    cleargrid grdInquiry
    FillGrid
    
    Dim vTitle             As String
    If optSalesPer.Value = True Then
        NEW_LogAudit "V", "Sales Executive Performance", "", "", "", "Sales Executive Performance by : " & cboSalesAE, "", ""
        vTitle = "Sales Executive Performance"
    ElseIf optVehStock.Value = True Then
        NEW_LogAudit "V", "VEHICLE ON STOCK", "", "", "", "VEHICLE ON STOCK:" & cboSalesAE, "", ""
        vTitle = "VEHICLE ON STOCK"
    ElseIf optCarRelease.Value = True Then
        NEW_LogAudit "V", "TOTAL RELEASED VEHICLES", "", "", "", "TOTAL RELEASED VEHICLES:" & cboModel & ":" & "Month:" & cboMonth, "", ""
        vTitle = "TOTAL RELEASED VEHICLES"
    ElseIf optInvCars.Value = True Then
        NEW_LogAudit "V", "INVOICED CARS", "", "", "", "INVOICED CARS:" & cboModel & ":" & "Month:" & cboMonth, "", ""
        vTitle = "INVOICED CARS"
    ElseIf optAllCars.Value = True Then
        NEW_LogAudit "V", "INVOICED CARS", "", "", "", "INVOICED CARS:" & cboModel & ":" & "Month:" & cboMonth, "", ""
    Else
        NEW_LogAudit "V", "SALES EXECUTIVE PERFORMANCE", "", "", "", "SA NAME: " & cboSalesAE & ":" & "Month:" & cboMonth, "", ""
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    '    rsMRRINV2.FILTER = ""
    '    FILTERVIEW
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If grdInquiry.Rows <= 2 Then Exit Sub
        If rsMRRINV2 Is Nothing Then Exit Sub
        grdInquiry.Height = 7305
        ShowHidePictureBox2 PICFILTER, True
    ElseIf KeyCode = vbKeyEscape And PICFILTER.Visible = True Then
        grdInquiry.Height = 8085
'        ShowHidePictureBox2 PICFILTER, False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If optAllCars.Value = True Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (ALLOCATED CARS)"
                Call frmALL_AuditInquiry.DisplayHistory("", "ALLOCATED CARS", "PRINTING")
            ElseIf optInvCars.Value = True Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (INVOICED CARS)"
                Call frmALL_AuditInquiry.DisplayHistory("", "INVOICED CARS", "PRINTING")
            ElseIf optCarRelease.Value = True Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TOTAL RELEASED VEHICLES)"
                Call frmALL_AuditInquiry.DisplayHistory("", "TOTAL RELEASED VEHICLES", "PRINTING")
            ElseIf optVehStock.Value = True Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLES ON STOCK)"
                Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLES ON STOCK", "PRINTING")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES EXECUTIVE PERFORMANCE)"
                Call frmALL_AuditInquiry.DisplayHistory("", "SALES EXECUTIVE PERFORMANCE", "PRINTING")
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    txtTotal = 0
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    cboSearchBy.AddItem "LIKE", 0
    cboSearchBy.AddItem "LOOSE MATCH", 1
    cboSearchBy.AddItem "EXACT MATCH", 2
    cboSearchBy.ListIndex = 0

    FillCboModel
    initGrid
    cleargrid grdInquiry
    Screen.MousePointer = 0
End Sub

Private Sub grdInquiry_DblClick()

    If GCOL <> grdInquiry.Col Then
        If GSORT = 1 Then
            grdInquiry.Sort = 2
            GSORT = 1
        Else
            grdInquiry.Sort = 1
            GSORT = 2
        End If

        GCOL = grdInquiry.Col

    Else
        If GSORT = 1 Then
            grdInquiry.Sort = 1
            GSORT = 2
        Else
            grdInquiry.Sort = 2
            GSORT = 1
        End If
        GCOL = grdInquiry.Col
    End If

End Sub

Private Sub grdInquiry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If grdInquiry.Row = 1 Then
        grdInquiry.MousePointer = flexCustom
    Else
        grdInquiry.MousePointer = flexArrow
    End If
End Sub

Private Sub optAllCars_Click()
    cleargrid grdInquiry
    initGrid
End Sub

Private Sub optCarRelease_Click()
    cleargrid grdInquiry
    initGrid
End Sub

Private Sub optInvCars_Click()
    cleargrid grdInquiry
    initGrid
End Sub

Private Sub Option1_Click()
    txtFilter_CS.Enabled = True: txtFilter_VI.Enabled = False
    txtFilter_CS.SetFocus
End Sub

Private Sub Option2_Click()
    txtFilter_CS.Enabled = False: txtFilter_VI.Enabled = True
    txtFilter_VI.SetFocus
End Sub

Private Sub optSalesPer_Click()
    cleargrid grdInquiry
    initGrid
End Sub

Private Sub optVehStock_Click()
    cleargrid grdInquiry
    initGrid
End Sub

Private Sub txtFilter_CS_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If KeyAscii = 13 Then
'        rsMRRINV2.Requery
        SHOWFILTERVIEW
    End If
End Sub

Private Sub txtFilter_VI_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If KeyAscii = 13 Then
        rsMRRINV2.Requery
        SHOWFILTERVIEW
    End If
End Sub

