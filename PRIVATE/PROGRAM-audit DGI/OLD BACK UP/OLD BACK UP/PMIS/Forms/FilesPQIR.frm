VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPMISTrans_PQIR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PQIR Form - Data Input"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FilesPQIR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11025
   ScaleWidth      =   11970
   Begin VB.PictureBox picBottom 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11970
      TabIndex        =   25
      Top             =   10125
      Width           =   11970
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   10050
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   35
         Top             =   30
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   750
            MouseIcon       =   "FilesPQIR.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   -30
            MouseIcon       =   "FilesPQIR.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Save this Record"
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   4680
         ScaleHeight     =   945
         ScaleWidth      =   6900
         TabIndex        =   26
         Top             =   0
         Width           =   6900
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   6120
            MouseIcon       =   "FilesPQIR.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Exit Window"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   5340
            MouseIcon       =   "FilesPQIR.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Print this Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   4560
            MouseIcon       =   "FilesPQIR.frx":1B6C
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":1CBE
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Delete Selected Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3780
            MouseIcon       =   "FilesPQIR.frx":1FE9
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":213B
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   3000
            MouseIcon       =   "FilesPQIR.frx":2497
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":25E9
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Add Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   2220
            MouseIcon       =   "FilesPQIR.frx":28FC
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":2A4E
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Find a Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1440
            MouseIcon       =   "FilesPQIR.frx":2D48
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":2E9A
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Width           =   795
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   660
            MouseIcon       =   "FilesPQIR.frx":31F2
            MousePointer    =   99  'Custom
            Picture         =   "FilesPQIR.frx":3344
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox picMiddles 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10125
      Left            =   0
      ScaleHeight     =   10125
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   0
      Width           =   11970
      Begin VB.VScrollBar ScrollBar1 
         Height          =   4965
         LargeChange     =   500
         Left            =   11610
         Max             =   11160
         SmallChange     =   250
         TabIndex        =   24
         Top             =   0
         Value           =   10
         Width           =   300
      End
      Begin VB.PictureBox picPQIR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   10500
         Left            =   -210
         Picture         =   "FilesPQIR.frx":36A3
         ScaleHeight     =   10470
         ScaleWidth      =   11865
         TabIndex        =   1
         Top             =   -390
         Width           =   11895
         Begin VB.TextBox TXTPONO 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4500
            TabIndex        =   49
            Top             =   1170
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtRRNO 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4500
            TabIndex        =   48
            Top             =   1590
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboDON 
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
            Left            =   450
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1920
            Width           =   2715
         End
         Begin VB.ComboBox cboPartNo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   7050
            TabIndex        =   46
            Text            =   "cboPartNo"
            Top             =   1470
            Width           =   4485
         End
         Begin VB.TextBox txtPART_NAME 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7050
            TabIndex        =   45
            Top             =   1830
            Width           =   4485
         End
         Begin VB.TextBox txtPQI_Code 
            Height          =   375
            Left            =   2670
            TabIndex        =   43
            Top             =   630
            Width           =   2745
         End
         Begin VB.CommandButton cmdEditTranDate 
            Caption         =   "..."
            Height          =   315
            Left            =   11280
            TabIndex        =   40
            Top             =   1110
            Width           =   255
         End
         Begin VB.PictureBox Picture1_ 
            Height          =   705
            Left            =   5850
            ScaleHeight     =   645
            ScaleWidth      =   105
            TabIndex        =   38
            Top             =   3390
            Visible         =   0   'False
            Width           =   165
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4590
            Top             =   1470
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin VB.TextBox txtPQINo 
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
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   -330
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add Photo"
            Height          =   375
            Left            =   10560
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1065
         End
         Begin VB.TextBox txtDescription 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   420
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   3150
            Width           =   5235
         End
         Begin VB.TextBox txtDiagnosis 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   420
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   4650
            Width           =   5235
         End
         Begin VB.TextBox txtRecommendation 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   420
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   7110
            Width           =   5235
         End
         Begin VB.TextBox txtSubject 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   7050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2220
            Width           =   2475
         End
         Begin VB.TextBox txtDate 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8700
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1110
            Width           =   2535
         End
         Begin VB.TextBox txtClaimType 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10830
            MaxLength       =   2
            TabIndex        =   6
            Top             =   2550
            Width           =   975
         End
         Begin VB.TextBox txtPART_NUMBER 
            BackColor       =   &H00E0E0E0&
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
            Left            =   420
            MaxLength       =   45
            TabIndex        =   14
            Top             =   9330
            Width           =   2145
         End
         Begin VB.TextBox txtPartDescription 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   15
            Top             =   9330
            Width           =   3075
         End
         Begin VB.TextBox txtDealerOrderNo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   5730
            MaxLength       =   45
            TabIndex        =   16
            Top             =   9330
            Width           =   2265
         End
         Begin VB.TextBox txtDealerSINO 
            BackColor       =   &H00E0E0E0&
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
            Left            =   390
            MaxLength       =   45
            TabIndex        =   19
            Top             =   9990
            Width           =   2175
         End
         Begin VB.TextBox txtDeliveryRecNo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   20
            Top             =   9990
            Width           =   3045
         End
         Begin VB.TextBox txtQty 
            BackColor       =   &H00E0E0E0&
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
            Left            =   5730
            MaxLength       =   10
            TabIndex        =   21
            Top             =   9990
            Width           =   2265
         End
         Begin VB.TextBox txtDateOrdered 
            BackColor       =   &H00E0E0E0&
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
            Left            =   8040
            MaxLength       =   20
            TabIndex        =   17
            Top             =   9330
            Width           =   1875
         End
         Begin VB.TextBox txtUnitAmount 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   22
            Top             =   9990
            Width           =   1875
         End
         Begin VB.TextBox txtDateReceived 
            BackColor       =   &H00E0E0E0&
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
            Left            =   9930
            MaxLength       =   20
            TabIndex        =   18
            Top             =   9330
            Width           =   1875
         End
         Begin VB.TextBox txtTotalAmount 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   9990
            Width           =   1845
         End
         Begin VB.TextBox txtSpecificInformation 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   450
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   8220
            Width           =   11295
         End
         Begin VB.TextBox txtReportedBy 
            BackColor       =   &H00E0E0E0&
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
            Left            =   390
            TabIndex        =   7
            Top             =   2580
            Width           =   2175
         End
         Begin VB.TextBox txtNotedBy 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2580
            TabIndex        =   8
            Top             =   2580
            Width           =   3105
         End
         Begin VB.Image Picture1 
            Height          =   4395
            Left            =   5850
            Stretch         =   -1  'True
            Top             =   3330
            Width           =   5865
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select PO No (HARI'S FORMAT)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   44
            Top             =   1620
            Width           =   1905
         End
         Begin VB.Label lbldate 
            Caption         =   "Label1"
            Height          =   345
            Left            =   10380
            TabIndex        =   42
            Top             =   -270
            Width           =   1275
         End
         Begin VB.Label lblDealer 
            Caption         =   "Label1"
            Height          =   315
            Left            =   8730
            TabIndex        =   41
            Top             =   -240
            Width           =   1515
         End
         Begin VB.Label lblPicPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6000
            TabIndex        =   39
            Top             =   -300
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label labID 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5340
            TabIndex        =   2
            Top             =   -300
            Visible         =   0   'False
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "frmPMISTrans_PQIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPQIR                                             As ADODB.Recordset
Dim ADDOREDIT                                          As String
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet
Dim fso                                                As FileSystemObject

Function GenerateCode(TABLENAME, FLDNAME As String, xFormat As String) As String
    Dim rsID                                           As ADODB.Recordset

    Set rsID = gconDMIS.Execute("Select MAX( ISNULL(" & FLDNAME & ", 0) ) as IDFIELD from " & TABLENAME)
    If rsID.Fields(0).Value = 0 Then
        GenerateCode = Format(1, xFormat)
    Else
        GenerateCode = Format(Val(N2Str2Zero(rsID![IDFIELD])) + 1, xFormat)
    End If
    Set rsID = Nothing

End Function

Sub fillPO()
    Combo_Loadval cboDON, gconDMIS.Execute("SELECT DISTINCT DON FROM PMIS_PO_HD WHERE DON IS NOT NULL UNION  SELECT DISTINCT DON FROM PMIS_PO_HIST WHERE DON IS NOT NULL ORDER BY DON")
End Sub

Sub rsRefresh()
    Set rsPQIR = New ADODB.Recordset
    Call rsPQIR.Open("SELECT  * FROM PMIS_PQIR order by id asc", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub InitMemVars()
    Dim cntrl                                          As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then

            If TypeOf cntrl Is ComboBox Then
                If cntrl.Style = 2 Then
                    cntrl.ListIndex = -1
                Else
                    cntrl.Text = vbNullString
                End If
            Else
                cntrl.Text = vbNullString
            End If

        End If
    Next
    txtDate = FormatDateTime(Now, vbShortDate)
    txtPQINo = GenerateCode("PMIS_PQIR", "PQINO", "0000")
    If CStr(Trim(lbldate)) <> "Label1" Then
        txtPQI_Code = lblDealer + "-" + CStr(Right(Year(lbldate), 2)) + "-" + CStr(Month(lbldate)) + "-" + txtPQINo
        txtQty = 1
        Set Picture1.Picture = Nothing
    End If
End Sub

Sub StoreMemvars()
    Dim mPath                                          As String
    labID = 0
    If Not (rsPQIR.EOF Or rsPQIR.BOF) Then
        With rsPQIR
            labID = Null2String(.Fields("ID"))
            txtClaimType = Null2String(.Fields("CLAIM_TYPE"))
            txtDate = Null2String(.Fields("DATEPQI"))
            txtDateOrdered = Null2String(.Fields("DATE_ORDERED"))
            txtDateReceived = Null2String(.Fields("DATE_RECEIVED"))
            txtDealerSINO = Null2String(.Fields("DEALER_SINO"))
            txtDealerOrderNo = Null2String(.Fields("DEALER_ORDER_NO"))
            txtDeliveryRecNo = Null2String(.Fields("DELIVERY_RECEIPT_NO"))
            txtDescription = Null2String(.Fields("DESCIPTIONS"))
            txtDiagnosis = Null2String(.Fields("DIAGNOSIS"))
            txtNotedBy = Null2String(.Fields("NOTEDBY"))
            txtPartDescription = Null2String(.Fields("PART_DESCRIPTION"))
            cboPartNo = Null2String(.Fields("PART_NO"))
            txtPART_NAME = Null2String(.Fields("PART_NAME"))
            txtPART_NUMBER = Null2String(.Fields("PART_NUMBER"))
            txtPQINo = Null2String(.Fields("PQINO"))
            txtQty = Null2String(.Fields("QUANTITY"))
            txtRecommendation = Null2String(.Fields("RECOMMENDATION"))
            txtReportedBy = Null2String(.Fields("REPORTEDBY"))
            txtSpecificInformation = Null2String(.Fields("SPECIFIC_INFO"))
            txtSubject = Null2String(.Fields("SUBJECT"))
            txtUnitAmount = Null2String(.Fields("UNITAMOUNT"))
            mPath = Null2String(.Fields("PICPATH"))
            lblPicPath = Null2String(.Fields("PICPATH"))

            txtPQI_Code = Null2String(.Fields("PQI_CODE"))

            If fso.FileExists(mPath) = True Then
                Picture1.Picture = LoadPicture(mPath)
            End If

        End With
    Else
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateAmount()
    txtTotalAmount = FormatNumber(NumericVal(txtQty) * NumericVal(txtUnitAmount))
End Sub

Private Sub cboPartNo_Change()
    cboPartNo_Click
End Sub

Private Sub cboPartNo_Click()
    If ADDOREDIT = "" Then Exit Sub
    Dim RSRR                                           As ADODB.Recordset
    Dim TEMPRS                                         As ADODB.Recordset

    txtPART_NAME = ""
    txtPART_NUMBER = ""
    txtPartDescription = ""
    txtUnitAmount = "0.00"
    Set TEMPRS = gconDMIS.Execute("SELECT  SRP,  STOCKDESC , STOCKNO FROM PMIS_STOCKMAS WHERE STOCKNO=" & N2Str2Null(cboPartNo))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtPART_NAME.Text = Null2String(TEMPRS!STOCKDESC)
        txtPART_NUMBER = Null2String(TEMPRS!STOCKNO)
        txtPartDescription = Null2String(TEMPRS!STOCKDESC)
        Set RSRR = gconDMIS.Execute("SELECT TRANINVAMT FROM PMIS_ALLDAYTRAN WHERE TRANTYPE='RR' AND TYPE='P' AND TRANNO=" & N2Str2Null(txtRRNO))
        If Not (RSRR.EOF Or RSRR.BOF) Then
            txtUnitAmount = FormatNumber(NumericVal(RSRR!TRANINVAMT))
        Else
            Set RSRR = gconDMIS.Execute("SELECT TRANINVAMT FROM PMIS_ALLDAYTRAN WHERE TRANTYPE='PO' AND TYPE='P' AND TRANNO=" & N2Str2Null(TXTPONO))
            If Not (RSRR.EOF Or RSRR.BOF) Then
                txtUnitAmount = FormatNumber(NumericVal(RSRR!TRANINVAMT))
            Else
                txtUnitAmount = FormatNumber(NumericVal(TEMPRS!SRP) / 1.32)
            End If
        End If

    End If

End Sub

Private Sub cboPartNo_GotFocus()
    If cboPartNo.Text = "<SELECT PART NUMBER FROM LIST>" Then
        cboPartNo.Text = ""
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "PARTS QUALITY INFORMATION REPORT") = False Then Exit Sub

    InitMemVars
    ADDOREDIT = "ADD": picAdds.Visible = False: picSaves.Visible = True: picPQIR.Enabled = True
    On Error Resume Next
    ' txtDate.SetFocus
End Sub

Private Sub cmdCancel_Click()
    ADDOREDIT = "": picAdds.Visible = True: picSaves.Visible = False: picPQIR.Enabled = False
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_delete", "PARTS QUALITY INFORMATION REPORT") = False Then Exit Sub
    On Error GoTo Errorcode:
    oVoice.Speak "Delete selected record? Are you sure?...", SVSFlagsAsync
    If MsgBox("Delete selected record, are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        If Not rsPQIR.EOF And Not rsPQIR.BOF Then
            Set rsPQIR = gconDMIS.Execute("DELETE from PMIS_PQIR where PQINO = '" & txtPQINo.Text & "'")
            ShowDeletedMsg
        End If
    End If
    LogAudit "X", "PQIR", txtPQINo
    rsRefresh
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_edit", "PARTS QUALITY INFORMATION REPORT") = False Then Exit Sub
    If NumericVal(labID) <> 0 Then
        ADDOREDIT = "EDIT"
        picAdds.Visible = False
        picSaves.Visible = True
        picPQIR.Enabled = True
    Else
        MessagePop RecSave, " Empty Record", "There is No Record To Delete"
    End If
End Sub

Private Sub cmdEditTranDate_Click()

    If Function_Access(LOGID, "Acess_SYSTEM", "PARTS QUALITY INFORMATION REPORT") = False Then Exit Sub
    txtDate.Enabled = True
    txtDate.Locked = False

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    rsPQIR.MoveNext
    If rsPQIR.EOF Then
        rsPQIR.MoveLast
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsPQIR.MovePrevious
    If rsPQIR.BOF Then
        rsPQIR.MoveFirst
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()

    If Function_Access(LOGID, "Acess_print", "PARTS QUALITY INFORMATION REPORT") = False Then Exit Sub

    If Len(Dir(App.Path & "\PQIR.xlt")) <= 0 Then
        If EXTRACT_FILES(107, "PQIR.xlt") = False Then
            MsgBox "Please Put PQIR.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    Screen.MousePointer = 11
    Dim vPQIRNo                                        As String
    Dim vPQIRDate                                      As String
    Dim vPQIRDateOrdered                               As String
    Dim vPQIRDateReceived                              As String
    Dim vPQIRPartNo                                    As String
    Dim vPQIRPartDesc                                  As String
    Dim vPQIRSubject                                   As String
    Dim vPQIRReportBy                                  As String
    Dim vPQIRNotedBy                                   As String
    Dim vPQIRClaimType                                 As String
    Dim vPQIRDescription                               As String
    Dim vPQIRDiagnosis                                 As String
    Dim vPQIRRecommendation                            As String
    Dim vPQIRSpecInfo                                  As String
    Dim vPQIRDetailsPartNo                             As String
    Dim vPQIRDetailsPartDesc                           As String
    Dim vPQIRDealerOrdNo                               As String
    Dim vPQIRDealerSINO                                As String
    Dim vPQIRDeliveryReceiptNo                         As String
    Dim vPQIRQty                                       As Integer
    Dim vPQIRUnitAmount                                As Double
    Dim vPQIRTotalAmount                               As Double
    Dim PQIR_CODE                                      As String

    Dim rsPQIR                                         As ADODB.Recordset

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "PQIR.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    Set rsPQIR = New ADODB.Recordset
    Set rsPQIR = gconDMIS.Execute("Select * from PMIS_PQIR where PQINO= '" & txtPQINo & "'")
    If Not rsPQIR.EOF And Not rsPQIR.BOF Then

        vPQIRNo = Null2String(rsPQIR!PQINO)
        vPQIRDate = Null2String(rsPQIR!DATEPQI)
        vPQIRDateOrdered = Null2String(rsPQIR!DATE_ORDERED)
        vPQIRDateReceived = Null2String(rsPQIR!Date_Received)
        vPQIRPartNo = Null2String(rsPQIR!PART_NO)
        vPQIRPartDesc = Null2String(rsPQIR!PART_DESCRIPTION)
        vPQIRSubject = Null2String(rsPQIR!Subject)
        vPQIRReportBy = Null2String(rsPQIR!REPORTEDBY)
        vPQIRNotedBy = Null2String(rsPQIR!NotedBy)
        vPQIRClaimType = Null2String(rsPQIR!CLAIM_TYPE)
        vPQIRDescription = Null2String(rsPQIR!DESCIPTIONS)
        vPQIRDiagnosis = Null2String(rsPQIR!Diagnosis)
        vPQIRRecommendation = Null2String(rsPQIR!RECOMMENDATION)
        vPQIRSpecInfo = Null2String(rsPQIR!SPECIFIC_INFO)
        vPQIRDetailsPartNo = Null2String(rsPQIR!PART_NUMBER)
        vPQIRDetailsPartDesc = Null2String(rsPQIR!PART_NAME)
        vPQIRDealerOrdNo = Null2String(rsPQIR!DEALER_ORDER_NO)
        vPQIRDealerSINO = Null2String(rsPQIR!DEALER_SINO)
        vPQIRDeliveryReceiptNo = Null2String(rsPQIR!DELIVERY_RECEIPT_NO)
        vPQIRQty = Null2String(rsPQIR!QUANTITY)
        vPQIRUnitAmount = Null2String(rsPQIR!UNITAMOUNT)
        vPQIRTotalAmount = vPQIRQty * vPQIRUnitAmount
        PQIR_CODE = Null2String(rsPQIR!PQI_CODE)

        'xlSheet.Cells(2, 4) = vPQIRNo
        xlSheet.Cells(2, 4) = PQIR_CODE
        xlSheet.Cells(4, 11) = vPQIRDate
        xlSheet.Cells(6, 9) = vPQIRPartNo
        xlSheet.Cells(7, 9) = vPQIRPartDesc
        xlSheet.Cells(10, 2) = vPQIRReportBy
        xlSheet.Cells(10, 5) = vPQIRNotedBy
        xlSheet.Cells(8, 9) = vPQIRSubject
        xlSheet.Cells(10, 13) = vPQIRClaimType
        xlSheet.Cells(13, 2) = vPQIRDescription
        xlSheet.Cells(13, 8) = vPQIRPartPix
        xlSheet.Cells(19, 2) = vPQIRDiagnosis
        xlSheet.Cells(29, 2) = vPQIRRecommendation
        xlSheet.Cells(39, 1) = vPQIRDetailsPartNo
        xlSheet.Cells(39, 4) = vPQIRDetailsPartDesc
        xlSheet.Cells(39, 8) = vPQIRDealerOrdNo
        xlSheet.Cells(39, 10) = vPQIRDateOrdered
        xlSheet.Cells(39, 12) = vPQIRDateReceived
        xlSheet.Cells(42, 1) = vPQIRDealerSINO
        xlSheet.Cells(42, 4) = vPQIRDeliveryReceiptNo
        xlSheet.Cells(42, 8) = vPQIRQty
        xlSheet.Cells(42, 10) = vPQIRUnitAmount
        xlSheet.Cells(42, 12) = vPQIRTotalAmount

        If Len(Dir(Null2String(rsPQIR!picpath))) > 0 Then
            Call xlSheet.Shapes.AddPicture(rsPQIR!picpath, 1, 1, 265, 150, 280, 230)
        End If
        xlApp.Visible = True
        Set xlApp = Nothing
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSave_Click()

    If LTrim(RTrim(txtClaimType)) = "" Then
        ShowIsRequiredMsg "Valid Claim Type"
        Exit Sub
    End If





    On Error GoTo Errorcode:

    If ADDOREDIT = "ADD" Then
        SQL = "INSERT INTO PMIS_PQIR("
        SQL = SQL & "PQINO, DATEPQI, PART_NO, PART_NAME, "
        SQL = SQL & "SUBJECT, CLAIM_TYPE, DESCIPTIONS, "
        SQL = SQL & "DIAGNOSIS, RECOMMENDATION, SPECIFIC_INFO, "
        SQL = SQL & "PART_NUMBER, PART_DESCRIPTION, DEALER_ORDER_NO, "
        SQL = SQL & "DATE_ORDERED, DATE_RECEIVED, DEALER_SINO, "
        SQL = SQL & "DELIVERY_RECEIPT_NO, QUANTITY, UNITAMOUNT, "
        SQL = SQL & "REPORTEDBY, NOTEDBY,PICPATH,PQI_CODE)"
        SQL = SQL & "VALUES("
        SQL = SQL & "@PQINO, @DatePQI, @PART_NO, @PART_NAME, "
        SQL = SQL & "@SUBJECT, @CLAIM_TYPE, @DESCRIPTIONS, "
        SQL = SQL & "@DIAGNOSIS, @RECOMMENDATION, @SPECIFIC_INFO, "
        SQL = SQL & "@PART_NUMBER, @PART_DESCRIPTION, @DEALER_ORDER_NO, "
        SQL = SQL & "@DATE_ORDERED, @DATE_RECEIVED, @DEALER_SINO, @DELIVERY_RECEIPT_NO, "
        SQL = SQL & "@QUANTITY, @UNITAMOUNT, @REPORTEDBY, @NOTEDBY,@PICPATH,@PQI_CODE)"
        LogAudit "A", "PQIR", txtPQINo
    Else
        SQL = "UPDATE PMIS_PQIR"
        SQL = SQL & " SET "
        SQL = SQL & " PQINO=@PQINO, "
        SQL = SQL & " DATEPQI=@DatePQI, "
        SQL = SQL & " PART_NO=@PART_NO, "
        SQL = SQL & " PART_NAME=@PART_NAME, "
        SQL = SQL & " SUBJECT=@SUBJECT, "
        SQL = SQL & " CLAIM_TYPE=@CLAIM_TYPE, "
        SQL = SQL & " DESCIPTIONS=@DESCRIPTIONS, "
        SQL = SQL & " DIAGNOSIS=@DIAGNOSIS, "
        SQL = SQL & " RECOMMENDATION=@RECOMMENDATION, "
        SQL = SQL & " SPECIFIC_INFO=@SPECIFIC_INFO, "
        SQL = SQL & " PART_NUMBER=@PART_NUMBER, "
        SQL = SQL & " PART_DESCRIPTION=@PART_DESCRIPTION, "
        SQL = SQL & " DEALER_ORDER_NO=@DEALER_ORDER_NO, "
        SQL = SQL & " DATE_ORDERED=@DATE_ORDERED, "
        SQL = SQL & " DATE_RECEIVED=@DATE_RECEIVED, "
        SQL = SQL & " DEALER_SINO=@DEALER_SINO, "
        SQL = SQL & " DELIVERY_RECEIPT_NO=@DELIVERY_RECEIPT_NO, "
        SQL = SQL & " QUANTITY=@QUANTITY, "
        SQL = SQL & " UNITAMOUNT=@UNITAMOUNT, "
        SQL = SQL & " REPORTEDBY=@REPORTEDBY, "
        SQL = SQL & " NOTEDBY=@NOTEDBY, "
        SQL = SQL & " PQI_CODE=@PQI_CODE,"
        SQL = SQL & " PICPATH=@PICPATH"
        SQL = SQL & " WHERE  ID=@ID"
        LogAudit "E", "PQIR", txtPQINo
    End If

    SQL = Replace(SQL, "@PQINO", N2Str2Null(txtPQINo))
    SQL = Replace(SQL, "@DatePQI", N2Str2Null(txtDate))
    SQL = Replace(SQL, "@PART_NO", N2Str2Null(cboPartNo))
    SQL = Replace(SQL, "@PART_NAME", N2Str2Null(txtPART_NAME))
    SQL = Replace(SQL, "@SUBJECT", N2Str2Null(txtSubject))
    SQL = Replace(SQL, "@CLAIM_TYPE", N2Str2Null(txtClaimType))
    SQL = Replace(SQL, "@DESCRIPTIONS", N2Str2Null(txtDescription))
    SQL = Replace(SQL, "@DIAGNOSIS", N2Str2Null(txtDiagnosis))
    SQL = Replace(SQL, "@RECOMMENDATION", N2Str2Null(txtRecommendation))
    SQL = Replace(SQL, "@SPECIFIC_INFO", N2Str2Null(txtSpecificInformation))
    SQL = Replace(SQL, "@PART_NUMBER", N2Str2Null(txtPART_NUMBER))
    SQL = Replace(SQL, "@PART_DESCRIPTION", N2Str2Null(txtPartDescription))
    SQL = Replace(SQL, "@DEALER_ORDER_NO", N2Str2Null(txtDealerOrderNo))
    SQL = Replace(SQL, "@DATE_ORDERED", N2Str2Null(txtDateOrdered))
    SQL = Replace(SQL, "@DATE_RECEIVED", N2Str2Null(txtDateReceived))
    SQL = Replace(SQL, "@DEALER_SINO", N2Str2Null(txtDealerSINO))
    SQL = Replace(SQL, "@DELIVERY_RECEIPT_NO", N2Str2Null(txtDeliveryRecNo))
    SQL = Replace(SQL, "@QUANTITY", NumericVal(txtQty))
    SQL = Replace(SQL, "@UNITAMOUNT", CCur(NumericVal(txtUnitAmount)))
    SQL = Replace(SQL, "@REPORTEDBY", N2Str2Null(txtReportedBy))
    SQL = Replace(SQL, "@NOTEDBY", N2Str2Null(txtNotedBy))
    SQL = Replace(SQL, "@PICPATH", N2Str2Null(lblPicPath))
    SQL = Replace(SQL, "@PQI_CODE", N2Str2Null(txtPQI_Code.Text))
    SQL = Replace(SQL, "@ID", labID)

    gconDMIS.Execute SQL

    rsRefresh
    rsPQIR.Find ("PQINO='" & txtPQINo & "'")
    StoreMemvars
    cmdCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cboDON_Click()

    Dim RSRR                                           As ADODB.Recordset
    Dim rsPODetail                                     As ADODB.Recordset
    Dim RSSTOCK                                        As ADODB.Recordset

    txtDateOrdered = ""
    cboPartNo.Clear
    txtDealerOrderNo = ""
    txtDateReceived = ""
    txtRRNO = ""

    Set rsPODetail = gconDMIS.Execute("SELECT  PONO,  PODATE  FROM PMIS_PO_HIST  WHERE DON='" & Repleys(cboDON) & "' UNION  SELECT  PONO,  PODATE     FROM PMIS_PO_HD WHERE DON='" & Repleys(cboDON) & "'")
    If Not (rsPODetail.EOF Or rsPODetail.BOF) Then
        txtDateOrdered = Null2String(rsPODetail!PODATE)
        txtDealerOrderNo = cboDON.Text
        TXTPONO = Null2String(rsPODetail!PONO)
        Set RSRR = gconDMIS.Execute("SELECT INVNO, DRNO, RRNO, RRDATE FROM PMIS_REC_HIST WHERE PONO='" & Null2String(rsPODetail!PONO) & "' UNION SELECT INVNO, DRNO, RRNO, RRDATE FROM PMIS_RR_HD WHERE PONO=" & Null2String(rsPODetail!PONO))
        If Not (RSRR.EOF Or RSRR.BOF) Then
            txtDateReceived = Null2String(RSRR!RRDATE)
            txtRRNO = Null2String(RSRR!RRNO)
            'txtDealerSINO
            'txtDealerSINO
        End If
        Set RSSTOCK = gconDMIS.Execute("SELECT STOCK_SUP FROM PMIS_ALLDAYTRAN WHERE TRANTYPE='PO' AND TYPE='P' AND TRANNO=" & N2Str2Null(rsPODetail!PONO))
        If Not RSSTOCK.EOF Or Not RSSTOCK.BOF Then
            While Not RSSTOCK.EOF
                If IsNull(RSSTOCK!STOCK_SUP) = False Then
                    cboPartNo.AddItem RSSTOCK!STOCK_SUP

                End If
                RSSTOCK.MoveNext
            Wend
            cboPartNo.Text = "<SELECT PART NUMBER FROM LIST>"
        End If
    End If



End Sub

Private Sub Command1_Click()

    On Error GoTo ErrorAddPic

    Dim pic_path                                       As String
    CommonDialog1.FILENAME = ""
    pic_path = ""
    CommonDialog1.Filter = "Graphic Files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
    CommonDialog1.ShowOpen


    If CommonDialog1.FILENAME <> "" Then
        Picture1.Picture = LoadPicture(CommonDialog1.FILENAME)
        If fso.FolderExists(PMIS_REPORT_PATH & "images") = False Then
            fso.CreateFolder (PMIS_REPORT_PATH & "images")
            pic_path = PMIS_REPORT_PATH & "images\" & txtPQI_Code & "." & Right(CommonDialog1.FILENAME, 3)
            fso.CopyFile CommonDialog1.FILENAME, pic_path, True
            lblPicPath = pic_path
        Else
            pic_path = PMIS_REPORT_PATH & "images\" & txtPQI_Code & "." & Right(CommonDialog1.FILENAME, 3)
            fso.CopyFile CommonDialog1.FILENAME, pic_path, True
            lblPicPath = pic_path
        End If
    End If

ErrorAddPic:
    Exit Sub


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Me.Height = 8500
    Set fso = New FileSystemObject
    picMiddles.Height = Me.ScaleHeight - picBottom.Height
    ScrollBar1.Height = picMiddles.ScaleHeight - 15
    ScrollBar1.Max = Abs(picMiddles.ScaleHeight - picPQIR.Height) + 20
    CenterMe frmMain, Me, 1
    fillPO
    InitMemVars
    rsRefresh
    StoreMemvars
    lblDealer.Caption = COMPANY_CODE
    lbldate.Caption = LOGDATE
    cmdCancel.Value = True
End Sub

Private Sub ScrollBar1_Change()
    picPQIR.Top = 0 - ScrollBar1.Value
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If IsDate(txtDate) = True Then
        txtDate = Format(txtDate, "mm/dd/yyyy")
    Else
        txtDate = Format(Now, "mm/dd/yyyy")
    End If
End Sub

Private Sub txtDateOrdered_Validate(Cancel As Boolean)
    If IsDate(txtDateOrdered) = True Then
        txtDateOrdered = Format(txtDateOrdered, "mm/dd/yyyy")
    Else
        txtDateOrdered = vbNullString
    End If
End Sub

Private Sub txtDateReceived_Validate(Cancel As Boolean)
    If IsDate(txtDateReceived) = True Then
        txtDateReceived = Format(txtDateReceived, "mm/dd/yyyy")
    Else
        txtDateReceived = vbNullString
    End If
End Sub

Private Sub txtQty_Change()
    UpdateAmount
End Sub

Private Sub txtUnitAmount_Change()
    UpdateAmount
End Sub

