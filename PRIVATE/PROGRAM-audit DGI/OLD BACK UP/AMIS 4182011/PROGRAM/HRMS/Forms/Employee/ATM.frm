VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMSATM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ATM Entry"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "ATM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8685
   Begin VB.PictureBox picEdit 
      Height          =   1185
      Left            =   6690
      ScaleHeight     =   1125
      ScaleWidth      =   1455
      TabIndex        =   22
      Top             =   1980
      Width           =   1515
      Begin VB.TextBox txtID 
         Height          =   345
         Left            =   180
         TabIndex        =   29
         Top             =   2580
         Width           =   1155
      End
      Begin VB.TextBox txtyear 
         Height          =   345
         Left            =   150
         TabIndex        =   28
         Top             =   2160
         Width           =   1155
      End
      Begin VB.TextBox txtmonth 
         Height          =   345
         Left            =   150
         TabIndex        =   27
         Top             =   1770
         Width           =   1155
      End
      Begin VB.TextBox txtcutoff 
         Height          =   345
         Left            =   150
         TabIndex        =   26
         Top             =   1320
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   24
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   90
         TabIndex        =   25
         Top             =   330
         Width           =   1275
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1455
         _Version        =   655364
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Edit ATM Entry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorDark=   -2147483635
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2730
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   10
      Top             =   3885
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "ATM.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4170
         MouseIcon       =   "ATM.frx":07C2
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Height          =   795
         Left            =   3480
         MouseIcon       =   "ATM.frx":0C7A
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   2790
         MouseIcon       =   "ATM.frx":10F7
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":1249
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Height          =   795
         Left            =   2100
         MouseIcon       =   "ATM.frx":15A5
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":16F7
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   1410
         MouseIcon       =   "ATM.frx":1A0A
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":1B5C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   720
         MouseIcon       =   "ATM.frx":1E56
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":1FA8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Next Record"
         Top             =   30
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
         Left            =   30
         MouseIcon       =   "ATM.frx":2300
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":2452
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picATM 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2580
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtYTDIncome 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4260
         TabIndex        =   32
         Top             =   510
         Width           =   1605
      End
      Begin VB.TextBox txtAccountNo 
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
         Left            =   1650
         TabIndex        =   30
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtPosition 
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
         Left            =   7020
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtName 
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
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   5775
      End
      Begin Crystal.CrystalReport rptATM 
         Left            =   7020
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Income"
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
         Height          =   315
         Left            =   3090
         TabIndex        =   33
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ATM Account No."
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
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label labID 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   5
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Height          =   315
         Left            =   6450
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdATM 
      Height          =   2535
      Left            =   2790
      TabIndex        =   4
      Top             =   1230
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   6
      ForeColor       =   0
      ForeColorFixed  =   0
      BackColorSel    =   14606302
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ATM.frx":27B1
   End
   Begin VB.PictureBox Picture11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   30
      Picture         =   "ATM.frx":2ACB
      ScaleHeight     =   4080
      ScaleWidth      =   2445
      TabIndex        =   6
      Top             =   570
      Width           =   2505
   End
   Begin VB.PictureBox picSearch 
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
      Height          =   4515
      Left            =   30
      Picture         =   "ATM.frx":16828
      ScaleHeight     =   4485
      ScaleWidth      =   2475
      TabIndex        =   7
      Top             =   180
      Width           =   2505
      Begin VB.TextBox txtSearch 
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   0
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   4125
         Left            =   30
         TabIndex        =   9
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   7276
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
         MouseIcon       =   "ATM.frx":19564
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
         Picture         =   "ATM.frx":196C6
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6870
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   19
      Top             =   3885
      Width           =   1440
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
         Height          =   795
         Left            =   720
         MouseIcon       =   "ATM.frx":2D433
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":2D585
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   30
         MouseIcon       =   "ATM.frx":2D8C3
         MousePointer    =   99  'Custom
         Picture         =   "ATM.frx":2DA15
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSATM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsATMdet                                               As ADODB.Recordset
Attribute rsATMdet.VB_VarUserMemId = 1073938432
Dim fielddate                                                         As Date
Dim fieldAmount                                                       As Double

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT ID, ACCOUNTNO, EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPINFOEMPNO.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT ID, ACCOUNTNO, EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL FROM HRMS_EMPINFO where empno = '" & frmHRMSEmpInfo.labID.Caption & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT ID, ACCOUNTNO, EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL, RESIGNED FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL ORDER BY LASTNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub InitGrid()
    With grdATM
        .Rows = 2
        .ColWidth(0) = 1300
        .ColWidth(1) = 2200
        .ColWidth(2) = 1
        .ColWidth(3) = 1
        .ColWidth(4) = 1
        .ColWidth(5) = 1
        .Row = 0
        .Col = 0
        .Text = "Date"
        .Col = 1
        .Text = "Net Amount"
        .Col = 2
        .Text = "ID"
    End With
End Sub

Sub InitMemvars()
    txtName.Text = ""
    txtAccountNo.Text = ""
    txtPosition.Text = ""
End Sub

Sub StoreMemVars()
    On Error GoTo Errorcode
    Dim CNT                                                           As Integer
    Dim VYTDIncome                                                    As Double
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        labID.Caption = rsEmpInfo!ID
        txtAccountNo.Text = Null2String(rsEmpInfo!ACCOUNTNO)
        txtPosition.Text = Null2String(rsEmpInfo![Position])
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
        Set rsATMdet = New ADODB.Recordset
        rsATMdet.Open "SELECT * FROM HRMS_ATMDET WHERE ATMID = " & rsEmpInfo!ID & " ORDER BY PAY_YEAR DESC, PAY_MONTH DESC, CUT_OFF DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsATMdet.EOF And Not rsATMdet.BOF Then
            rsATMdet.MoveFirst
            cleargrid grdATM
            grdATM.Rows = grdATM.Rows
            CNT = 0
            VYTDIncome = 0
            Do While Not rsATMdet.EOF
                CNT = CNT + 1
                grdATM.AddItem Null2Date(rsATMdet!DEYT) & Chr(9) & N2Str2Zero(rsATMdet!netamount) & Chr(9) & rsATMdet!atmID & Chr(9) & rsATMdet!CUT_OFF & Chr(9) & rsATMdet!PAY_MONTH & Chr(9) & rsATMdet!PAY_YEAR
                If YEAR(Null2Date(rsATMdet!DEYT)) = YEAR(LOGDATE) Then
                    VYTDIncome = VYTDIncome + N2Str2Zero(rsATMdet!netamount)
                End If
                rsATMdet.MoveNext
            Loop
            If CNT > 0 Then grdATM.RemoveItem 1
        Else
            cleargrid grdATM
        End If
        txtYTDIncome.Text = N2Str2Zero(VYTDIncome)
    Else
        ShowNoRecord
        Unload Me
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub store_entry()
    Label4.Caption = grdATM.TextMatrix(grdATM.RowSel, 0)
    Text1.Text = grdATM.TextMatrix(grdATM.RowSel, 1)
    txtcutoff.Text = grdATM.TextMatrix(grdATM.RowSel, 3)
    txtmonth.Text = grdATM.TextMatrix(grdATM.RowSel, 4)
    txtyear.Text = grdATM.TextMatrix(grdATM.RowSel, 5)
    txtID.Text = grdATM.TextMatrix(grdATM.RowSel, 2)
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    picEdit.Enabled = False
    grdATM.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN ATM ENTRY") = False Then Exit Sub
    picEdit.Enabled = True
    grdATM.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error GoTo Errorcode:
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtSearch.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN ATM ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    rptATM.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptATM.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptATM.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptATM, HRMS_REPORT_PATH & "IndivATM.rpt", "{empinfo.empno} = " & N2Str2Null(rsEmpInfo!EMPNO), DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode:
    gconDMIS.Execute "UPDATE HRMS_ATMDET SET " & _
                   " NETAMOUNT = " & Text1.Text & _
                   " WHERE ATMID = " & txtID.Text & _
                   " AND CUT_OFF = " & txtcutoff.Text & _
                   " AND PAY_MONTH = " & txtmonth.Text & _
                   " AND PAY_YEAR = " & txtyear.Text
    ShowSuccessFullyUpdated
    cmdCancel.Value = True
    rsrefresh
    On Error Resume Next
    rsEmpInfo.Find "id = " & labID.Caption
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtSearch.Text = ""
    rsrefresh
    InitGrid
    InitMemvars
    StoreMemVars
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdATM_Click()
    store_entry
End Sub

Private Sub grdATM_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    rsEmpInfo.Bookmark = rsFind(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

