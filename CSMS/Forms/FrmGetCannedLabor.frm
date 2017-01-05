VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmCSMSGetCannedLabor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Canned Labor"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmGetCannedLabor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10785
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6285
      Left            =   3420
      ScaleHeight     =   6255
      ScaleWidth      =   7305
      TabIndex        =   9
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtnotes 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1260
         Width           =   7095
      End
      Begin VB.TextBox txtFlatrate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2850
         Width           =   1635
      End
      Begin VB.TextBox txtstdTime 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2850
         Width           =   1635
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   5805
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1245
      End
      Begin MSComctlLib.ListView lstJobDetails 
         Height          =   2805
         Left            =   120
         TabIndex        =   17
         Top             =   3300
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4948
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
         Appearance      =   1
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
         MouseIcon       =   "FrmGetCannedLabor.frx":058A
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code Header"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Job Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "STD Time"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Flat Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Notes/ Suggested Jobs"
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
         Left            =   150
         TabIndex        =   21
         Top             =   1050
         Width           =   1920
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Flat Rate"
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
         Left            =   1830
         TabIndex        =   20
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Standard Time"
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
         Left            =   150
         TabIndex        =   19
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Service Operation"
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
         Left            =   150
         TabIndex        =   18
         Top             =   390
         Width           =   1470
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   7425
         _Version        =   655364
         _ExtentX        =   13097
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Canned labor information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6285
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   3405
      TabIndex        =   6
      Top             =   0
      Width           =   3435
      Begin VB.TextBox txtKeyword 
         Height          =   360
         Left            =   60
         TabIndex        =   8
         Top             =   390
         Width           =   3315
      End
      Begin MSComctlLib.ListView lstCanned 
         Height          =   5295
         Left            =   60
         TabIndex        =   7
         Top             =   840
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   9340
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
         Appearance      =   1
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
         MouseIcon       =   "FrmGetCannedLabor.frx":06EC
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Canned Description"
            Object.Width           =   11465
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "STD Time"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Flat Rate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Canned Notes"
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3765
         _Version        =   655364
         _ExtentX        =   6641
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Search Canned labor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1875
      MouseIcon       =   "FrmGetCannedLabor.frx":084E
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetCannedLabor.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to Add/Edit/Delete Jobs"
      Top             =   6315
      Visible         =   0   'False
      Width           =   705
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
      Height          =   735
      Left            =   10020
      MouseIcon       =   "FrmGetCannedLabor.frx":0CB3
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetCannedLabor.frx":0E05
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   6330
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9300
      MouseIcon       =   "FrmGetCannedLabor.frx":1143
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetCannedLabor.frx":1295
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Select"
      Top             =   6330
      Width           =   735
   End
   Begin VB.Label LBLRO 
      BackColor       =   &H000000FF&
      Caption         =   "Label7"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   6330
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "&Add/Edit/Delete Jobs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2625
      TabIndex        =   4
      Top             =   6615
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label txtCheckMe 
      BackColor       =   &H000000FF&
      Caption         =   "Label7"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   6690
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmCSMSGetCannedLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUPLOAD                                           As ADODB.Recordset
Dim X                                                  As Long

Private Sub cmdAdd_Click()
    frmCSMSCannedlabor.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim ictr                                                    As Integer
    If txtCheckMe = "ro" Then
'        With frmCSMSNewAppointment.lblJob4Service
'            .Sorted = False
'            .ListItems.Add , , txtCode
'            .ListItems(.ListItems.Count).ListSubItems.Add 1, , "CND"
'            .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtDesc
'            .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtFlatrate
'            .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtstdTime
'            .ListItems(.ListItems.Count).ListSubItems.Add 5, , 0
'            .ListItems(.ListItems.Count).ListSubItems.Add 6, , ""
'            .ListItems(.ListItems.Count).ListSubItems.Add 7, , txtnotes
'        End With

        For X = 1 To lstJobDetails.ListItems.Count
            For ictr = 1 To frmCSMSNewAppointment.lblJob4Service.ListItems.Count
                If (lstJobDetails.ListItems(X).Text) = frmCSMSNewAppointment.lblJob4Service.ListItems(ictr) Then
                   MsgBox "Job already Added in this R.O." & vbCrLf & "" & N2Str2Null(lstJobDetails.ListItems(X).ListSubItems(2).Text) & " will not be added.", vbInformation + vbOKOnly, "Action Void"
                   GoTo next2
                End If
            Next ictr
            
            With frmCSMSNewAppointment.lblJob4Service
                .Sorted = False
                .ListItems.Add , , lstJobDetails.ListItems(X)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , lstJobDetails.ListItems(X).SubItems(5)
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , lstJobDetails.ListItems(X).SubItems(2)
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , lstJobDetails.ListItems(X).SubItems(4)
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , lstJobDetails.ListItems(X).SubItems(3)
                .ListItems(.ListItems.Count).ListSubItems.Add 5, , 0
                .ListItems(.ListItems.Count).ListSubItems.Add 6, , ""
                .ListItems(.ListItems.Count).ListSubItems.Add 7, , lstJobDetails.ListItems(X).SubItems(2)
            End With
            
'            With frmCSMSNewAppointment.lstPMSDet
'                .Sorted = False
'                .ListItems.Add , , lstJobDetails.ListItems(X)
'                .ListItems(.ListItems.Count).ListSubItems.Add 1, , "CND"
'                .ListItems(.ListItems.Count).ListSubItems.Add 2, , lstJobDetails.ListItems(X).SubItems(2)
'                .ListItems(.ListItems.Count).ListSubItems.Add 3, , lstJobDetails.ListItems(X).SubItems(1)
'            End With
next2:
        Next X
        
        Call cmdCancel_Click
        Exit Sub
    End If
    
    If txtCheckMe.Caption = "MAIN" Then
        Dim JOBREP_OR                                           As String
        Dim JOBLEVEL                                            As String
        Dim JOBLINE_NO                                          As String
        Dim JOBDETCDE                                           As String
        Dim VLastUpdateTime                                     As String
        Dim JOBDETDSC                                           As String
        Dim JOBDETUNT                                           As String
        Dim VLastUpdate                                         As String
        Dim Vusercode                                           As String
        Dim JOBDETVOL                                           As Double
        Dim JOBDETPRC                                           As Double
        Dim JOBDETAMT                                           As Double
        Dim JOBCODE                                             As String
        Dim JOBWCODE                                            As String
        Dim JOBTAXRATE                                          As Double
        Dim JOBDISCRATE                                         As Double
        Dim JOBTAXVAL                                           As Double
        Dim JOBDISVAL                                           As Double
        Dim JOBPOCODE                                           As String
        Dim JOBRep_Or2                                          As String
        Dim JOBDETAIL                                           As String
        Dim JOBDET_AMT                                          As Double
        Dim JOBDIS_VAL                                          As Double
        Dim JOBDISCOUNT_2                                       As Double
        Dim xFLATRATE                                           As Double
        Dim JOBREMARKS                                          As String
        Dim JOBTECHNICIAN                                       As String
        Dim JOBDET_HRS                                          As String
        Dim xJobType                                            As String
        Dim rstmp                                               As New ADODB.Recordset
        
        Dim xCNT                                                As Integer
        
        
        For xCNT = 1 To lstJobDetails.ListItems.Count
        'updated by: IEBV
        'description: To check if job's already added
            Dim RSDUP                           As New ADODB.Recordset
            Set RSDUP = gconDMIS.Execute("Select * from CSMS_RO_DET where detcde = " & N2Str2Null(lstJobDetails.ListItems(xCNT).Text) & " and rep_or = " & N2Str2Null(LBLRO.Caption) & " and transtype = 'R'")
            If Not (RSDUP.EOF And RSDUP.BOF) Then
                MsgBox "Job already Added in this R.O." & vbCrLf & "" & N2Str2Null(lstJobDetails.ListItems(xCNT).ListSubItems(2).Text) & " will not be added.", vbInformation + vbOKOnly, "Action Void"
                Set RSDUP = Nothing
                GoTo nextna
            End If
         '------------------------------------------------------
            Set rstmp = gconDMIS.Execute("SELECT LINE_NO FROM CSMS_RO_DET WHERE " & _
                " REP_OR = '" & LBLRO.Caption & _
                "' AND LIVIL = '1' ORDER BY LINE_NO DESC")
            If Not (rstmp.BOF And rstmp.EOF) Then
                JOBLINE_NO = Format(Val(rstmp!LINE_NO), "00")
            End If
    
            JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
            JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
    
            JOBREP_OR = N2Str2Null(LBLRO.Caption)
            JOBLEVEL = "'1'"
            JOBLINE_NO = N2Str2Null(Format(Val(JOBLINE_NO) + 1, "00"))
            JOBDETCDE = N2Str2Null(lstJobDetails.ListItems(xCNT).Text)
            xJobType = N2Str2Null(lstJobDetails.ListItems(xCNT).ListSubItems(5).Text)
            JOBDETDSC = N2Str2Null(lstJobDetails.ListItems(xCNT).ListSubItems(2).Text)
            xFLATRATE = NumericVal(lstJobDetails.ListItems(xCNT).ListSubItems(4).Text)
            JOBDET_HRS = NumericVal(lstJobDetails.ListItems(xCNT).ListSubItems(3).Text)
            JOBDISCRATE = NumericVal(0)
            JOBWCODE = "NULL"
            JOBDETUNT = "NULL"
            JOBDETVOL = NumericVal(0)
            JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
            JOBCODE = "NULL"
            JOBTAXRATE = (VAT_RATE / 100)
            JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
            JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    
            JOBRep_Or2 = "NULL"
            JOBDETAIL = N2Str2Null(lstJobDetails.ListItems(xCNT).ListSubItems(2).Text)
            JOBDET_AMT = JOBDETPRC
            JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
            JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
            JOBTECHNICIAN = "NULL"
            
            'COMMENT BY  : MJP 10162009 0336PM
            'DESCRIPTION : DOUBLE VAT
                'JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
            'COMMENT BY  : MJP 10162009 0336PM
            'UPDATE BY   : MJP 10162009 0336PM
                JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
            'UPDATE BY   : MJP 10162009 0336PM
            
            Vusercode = "" & N2Str2Null(LOGCODE) & ""
            VLastUpdate = "'" & LOGDATE & "'"
            VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    
            SQL_STATEMENT = "insert into CSMS_RO_Det (JobType, FLATRATE, rep_or, livil, LINE_NO, detcde, " & _
                "detdsc, det_hrs, detvol, detprc, detamt, wcode, taxrate, discrate, taxval, disval, detail, det_amt, dis_val, discount_2, USERCDE, SAVEDATE, SAVETIME) values (" & _
                xJobType & _
                ", " & xFLATRATE & _
                ", " & JOBREP_OR & _
                ", " & JOBLEVEL & _
                ", " & JOBLINE_NO & _
                ", " & JOBDETCDE & _
                ", " & JOBDETDSC & _
                ", " & JOBDET_HRS & _
                ", " & JOBDETVOL & _
                ", " & JOBDETPRC & _
                ", " & JOBDETAMT & _
                ", " & JOBWCODE & _
                ", " & (JOBTAXRATE * 100) & _
                ", " & (JOBDISCRATE * 100) & _
                ", " & JOBTAXVAL & _
                ", " & JOBDISVAL & _
                ", " & JOBDETAIL & _
                ", " & JOBDET_AMT & _
                ", " & JOBDIS_VAL & _
                ", " & JOBDISCOUNT_2 & _
                ", " & Vusercode & _
                ", " & VLastUpdate & _
                ", " & VLastUpdateTime & ")"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(JOBREP_OR), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & Null2String(JOBDETCDE), "", "")
            'NEW LOG AUDIT-----------------------------------------------------
nextna:
        Next
        Call CheckSTATUS
        frmCSMS_ServiceCounter.Click_ScheduleGrid (LBLRO.Caption)
        Call cmdCancel_Click
        Call ShowSuccessFullyAdded
    End If
End Sub
Sub CheckSTATUS()
    Dim RS                                             As New ADODB.Recordset
    Dim theRo                                          As String
    theRo = Trim(LBLRO.Caption)
    Set RS = gconDMIS.Execute("Select Case status " & _
                              " when 'Y' then 1 " & _
                              " when 'G' then 3 " & _
                              " when 'L' then 4 " & _
                              " when 'B' then 5 " & _
                              " when 'I' then 6 " & _
                              " when 'R' then 7 " & _
                              " else 2 " & _
                              " end as status1,status,rep_or from csms_ro_det where livil = 1 and rep_or = '" & theRo & "'order by status1 asc")
   If Not (RS.EOF And RS.BOF) Then
        RS.MoveFirst
        Do While Not RS.EOF
            If N2Str2Zero(RS!status1) = "1" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET jstatus='F', status='Finish Job' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "2" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = NULL, status='Park' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "3" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'G', status='Going Home' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "4" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'L', status='Lunch Break' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "5" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'B', status='Break Time' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "6" Then
                gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'I', status='Idle Time' where Ro_no='" & theRo & "'"
            End If
            RS.MoveNext
        Loop
   End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Call txtKeyword_Change
    
    Screen.MousePointer = 0
End Sub

Private Sub lstCanned_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstJobDetails.Enabled = False
    txtCode = lstCanned.SelectedItem
    txtDesc = lstCanned.SelectedItem.SubItems(1)
    txtstdTime = lstCanned.SelectedItem.SubItems(2)
    txtFlatrate = lstCanned.SelectedItem.SubItems(3)
    txtnotes = lstCanned.SelectedItem.SubItems(4)

    lstJobDetails.Sorted = False: lstJobDetails.ListItems.Clear
    Set RSUPLOAD = gconDMIS.Execute("select CODE,codeheader,Canned_Description,STDtime,FlatRate,jobtype from CSMS_CannedDetails where CODEHeader = '" & txtCode & "' order by Canned_Description asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstJobDetails.ListItems, RSUPLOAD
        lstJobDetails.Enabled = True
    End If

End Sub

Private Sub txtKeyword_Change()
    Set RSUPLOAD = New ADODB.Recordset
    lstCanned.Enabled = False
    lstCanned.Sorted = False: lstCanned.ListItems.Clear
    Set RSUPLOAD = gconDMIS.Execute("select CODE,Canned_Description,TimeSTD,FlatRate,CannedNotes from CSMS_CannedLabor " & _
        " where Canned_Description  like '" & txtKeyword & "%' " & _
        "order by Canned_Description asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstCanned.ListItems, RSUPLOAD
        lstCanned.Enabled = True
    End If
End Sub



