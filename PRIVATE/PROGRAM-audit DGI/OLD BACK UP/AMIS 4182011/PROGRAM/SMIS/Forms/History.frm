VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSMIS_Mis_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Call History"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "History.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture10 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10695
      TabIndex        =   5
      Top             =   975
      Width           =   10695
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   345
         Left            =   3570
         TabIndex        =   10
         Top             =   270
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   90
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   210
         Width           =   1335
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1470
         TabIndex        =   8
         Top             =   240
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52559873
         CurrentDate     =   39233
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3900
         TabIndex        =   9
         Top             =   210
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52559873
         CurrentDate     =   39233
      End
      Begin VB.Label Label7 
         Caption         =   "Date To"
         Height          =   285
         Left            =   3810
         TabIndex        =   12
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "Date From"
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   0
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture9 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10695
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      Begin VB.ComboBox cboCustomer 
         Height          =   345
         Left            =   210
         TabIndex        =   2
         Top             =   390
         Width           =   3405
      End
      Begin VB.Label labCustDetail 
         Caption         =   "Customer Name"
         Height          =   885
         Left            =   4890
         TabIndex        =   4
         Top             =   0
         Width           =   5670
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   120
         Width           =   1380
      End
   End
   Begin TabDlg.SSTab SearchTab 
      Height          =   7095
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   697
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Call History"
      TabPicture(0)   =   "History.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstSalesHistory"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vehicle Sales History"
      TabPicture(1)   =   "History.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vehicle Service History"
      TabPicture(2)   =   "History.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstServiceHistory"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Incident History"
      TabPicture(3)   =   "History.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstIncidentHistory"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin MSComctlLib.ListView lstIncidentHistory 
         Height          =   5025
         Left            =   -74940
         TabIndex        =   13
         Top             =   600
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   8864
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
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
         MouseIcon       =   "History.frx":037A
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstServiceHistory 
         Height          =   2535
         Left            =   -74970
         TabIndex        =   14
         Top             =   570
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
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
         MouseIcon       =   "History.frx":0694
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5235
         Left            =   60
         TabIndex        =   15
         Top             =   450
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   9234
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
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
         MouseIcon       =   "History.frx":09AE
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstSalesHistory 
         Height          =   5235
         Left            =   -74970
         TabIndex        =   16
         Top             =   450
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   9234
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
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
         MouseIcon       =   "History.frx":0CC8
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2955
         Left            =   -74940
         TabIndex        =   17
         Top             =   3150
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5212
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
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
         MouseIcon       =   "History.frx":0FE2
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmSMIS_Mis_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCust                             As Recordset
Dim WithEvents cCombo                  As cCombo
Attribute cCombo.VB_VarHelpID = -1
Private Sub cboCustomer_click()

'rsCust.Find("CUSCDE=" &)
End Sub

Private Sub cCombo_SelectionChange(Xl As Long, lXn As lAction)

    Dim ar                             As Variant
    ar = Split(cboCustomer.Text, ":")

    rsCust.MoveFirst
    rsCust.Find "CUSCDE='" & LTrim(RTrim(ar(1))) & "'"
    labCustDetail = " ContactPerson: " & Null2String(rsCust!ContactPerson) & vbCrLf & _
                  " Addres: " & Null2String(rsCust!Address)
End Sub

Private Sub Form_Load()
    Set rsCust = New ADODB.Recordset
    Set cCombo = New cCombo
    rsCust.Open "Select CustomerName +  ':' + CUSCDE as CUSTNAME, Address,CUSTID ,ContactPerson,CUSCDE  from CRIS_vw_allProfile Order By AcctName ", gconDMIS, adOpenKeyset, adLockReadOnly
    If Not (rsCust.EOF Or rsCust.BOF) Then
        While Not rsCust.EOF
            cboCustomer.AddItem Null2String(rsCust!CustName)
            rsCust.MoveNext
        Wend
        rsCust.MoveFirst
    End If


    Set cCombo.AttachCombo = cboCustomer
    SetComboWidth cboCustomer, 400
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    Select Case SearchTab.Tabs
        Case 0
        Case 1
        Case 2
        Case 3
    End Select
End Sub

