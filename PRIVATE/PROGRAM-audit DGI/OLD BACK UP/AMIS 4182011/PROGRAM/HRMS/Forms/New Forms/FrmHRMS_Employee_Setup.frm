VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form FrmHRMS_Employee_Setup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Payroll Setup"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "FrmHRMS_Employee_Setup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox piclst 
      Height          =   4905
      Left            =   60
      ScaleHeight     =   4845
      ScaleWidth      =   4755
      TabIndex        =   15
      Top             =   1320
      Width           =   4815
      Begin MSComctlLib.ListView lstemployee 
         Height          =   4905
         Left            =   -30
         TabIndex        =   16
         Top             =   -30
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp No"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6103
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3120
      ScaleHeight     =   945
      ScaleWidth      =   1740
      TabIndex        =   14
      Top             =   6300
      Width           =   1740
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
         Left            =   210
         MouseIcon       =   "FrmHRMS_Employee_Setup.frx":1BF62
         MousePointer    =   99  'Custom
         Picture         =   "FrmHRMS_Employee_Setup.frx":1C0B4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
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
         Left            =   930
         MouseIcon       =   "FrmHRMS_Employee_Setup.frx":1C410
         MousePointer    =   99  'Custom
         Picture         =   "FrmHRMS_Employee_Setup.frx":1C562
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.OptionButton optLN 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   990
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.OptionButton optempno 
      Caption         =   "Employee No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1590
      TabIndex        =   5
      Top             =   990
      Width           =   1665
   End
   Begin VB.CheckBox chk_check 
      Caption         =   "Select All"
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
      Height          =   525
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Select or Deselect Employee"
      Top             =   6270
      Width           =   1245
   End
   Begin VB.TextBox txtsearch 
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
      Height          =   345
      Left            =   690
      TabIndex        =   3
      Top             =   600
      Width           =   2745
   End
   Begin VB.TextBox txtyear 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   60
      Width           =   1125
   End
   Begin VB.TextBox txtmonth 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1830
      TabIndex        =   1
      Top             =   60
      Width           =   1425
   End
   Begin VB.TextBox txtcutoff 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3300
      ScaleHeight     =   945
      ScaleWidth      =   1590
      TabIndex        =   13
      Top             =   6270
      Width           =   1590
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
         Left            =   120
         MouseIcon       =   "FrmHRMS_Employee_Setup.frx":1C8C8
         MousePointer    =   99  'Custom
         Picture         =   "FrmHRMS_Employee_Setup.frx":1CA1A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save Entry"
         Top             =   30
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
         Height          =   795
         Left            =   810
         MouseIcon       =   "FrmHRMS_Employee_Setup.frx":1CD6A
         MousePointer    =   99  'Custom
         Picture         =   "FrmHRMS_Employee_Setup.frx":1CEBC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Search"
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
      Left            =   60
      TabIndex        =   12
      Top             =   660
      Width           =   570
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   525
      Left            =   0
      TabIndex        =   11
      Top             =   -30
      Width           =   5655
      _Version        =   655364
      _ExtentX        =   9975
      _ExtentY        =   926
      _StockProps     =   14
      ForeColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   14215660
   End
End
Attribute VB_Name = "FrmHRMS_Employee_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_check_Click()
    Dim all As Integer
       
    If chk_check.Value = 1 Then
        For all = 1 To lstemployee.ListItems.count
            lstemployee.ListItems(all).Checked = True
        Next
    Else
        For all = 1 To lstemployee.ListItems.count
            lstemployee.ListItems(all).Checked = False
        Next
    End If


End Sub

Private Sub cmdCancel_Click()

    Picture3.Visible = True
    Picture2.Visible = False
    piclst.Enabled = False
    txtsearch.Text = ""
    txtsearch.Enabled = False
    chk_check.Enabled = False
    optLN.Enabled = False
    optempno.Enabled = False
    
    FillSearchGrid ("")

End Sub

Private Sub cmdEdit_Click()

    Picture3.Visible = False
    Picture2.Visible = True
    piclst.Enabled = True
    txtsearch.Enabled = True
    chk_check.Enabled = True
    optLN.Enabled = True
    optempno.Enabled = True

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

Dim all As Integer
Dim SQL As String
Dim XEMPNO As String

For all = 1 To lstemployee.ListItems.count
     
        If lstemployee.ListItems(all).Checked = True Then
                    
            XEMPNO = Null2String(lstemployee.ListItems(all).Text)
            SQL = "UPDATE HRMS_EMPINFO SET EMPNO = " & XEMPNO & ",Includthispayroll = 'Y' where EMPNO= " & XEMPNO & " "
            gconDMIS.Execute (SQL)
            
        Else
        
            XEMPNO = Null2String(lstemployee.ListItems(all).Text)
            SQL = "UPDATE HRMS_EMPINFO SET EMPNO = " & XEMPNO & ",Includthispayroll = 'N' where EMPNO= " & XEMPNO & " "
            gconDMIS.Execute (SQL)
        
        End If
    
 Next

        ShowSuccessFullyUpdated
        Call cmdCancel_Click

End Sub

Private Sub Form_Load()
  
  Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Call detail_cutoff
  
    FillSearchGrid ("")
    
    optLN.Value = True
    piclst.Enabled = False
    txtsearch.Enabled = False
    chk_check.Enabled = False
    optLN.Enabled = False
    optempno.Enabled = False
 

 Screen.MousePointer = 0
 

End Sub

Sub detail_cutoff()

    Dim rsCutoff                                       As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            txtcutoff.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            txtcutoff.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        txtmonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        txtyear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If

End Sub


Sub Fillemployeegrid()
    
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT EMPNO, LASTNAME+', '+FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE activeinactive = 'A' ORDER BY LASTNAME asc")
    lstemployee.ListItems.Clear
    
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Call Listview_Loadval(lstemployee.ListItems, RSTMP)
    End If
    Set RSTMP = Nothing
End Sub

Sub FillGrid()

    Dim RSCUSTOMER                                     As ADODB.Recordset

    lstemployee.ListItems.Clear
    Set RSCUSTOMER = New ADODB.Recordset
    Set RSCUSTOMER = gconDMIS.Execute("SELECT EMPNO, LASTNAME+', '+FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE activeinactive = 'A' ORDER BY LASTNAME asc")
  

    If Not (RSCUSTOMER.EOF And RSCUSTOMER.BOF) Then
        Listview_Loadval Me.lstemployee.ListItems, RSCUSTOMER
        lstemployee.Enabled = True
        lstemployee.Refresh
    Else
        'do nothing
    End If

End Sub


Sub FillSearchGrid(XXX As String)

    Dim RSCUSTOMER                                     As ADODB.Recordset
    Dim SEARCHFILTER                                   As String
    Dim lst                                            As ListItem
    
    
    lstemployee.ListItems.Clear
    
    
    Set RSCUSTOMER = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    SEARCHFILTER = ""
    
    
        If optLN.Value = True Then
            Set RSCUSTOMER = gconDMIS.Execute("select empno, LASTNAME+', '+FIRSTNAME AS FULLNAME,* from hrms_empinfo where LastName like '" & XXX & "%'" & SEARCHFILTER & " order by lastname asc")
        Else: optempno.Value = True
            Set RSCUSTOMER = gconDMIS.Execute("select empno, LASTNAME+', '+FIRSTNAME AS FULLNAME,* from hrms_empinfo where empno like '" & XXX & "%'" & SEARCHFILTER & " order by middlename asc")
        End If
    
            
            While Not RSCUSTOMER.EOF
                Set lst = lstemployee.ListItems.Add(, , Null2String(RSCUSTOMER!EMPNO))
                 lst.ListSubItems.Add , , (Null2String(RSCUSTOMER!lastname) + ", " + Null2String(RSCUSTOMER!FIRSTNAME))
                If RSCUSTOMER!Includthispayroll = "N" Then
                    lst.Checked = False
                ElseIf RSCUSTOMER!Includthispayroll = "Y" Then
                    lst.Checked = True
        
                    lst.ForeColor = vbBlue
                    lst.Bold = True
        
                    lst.ListSubItems(1).ForeColor = vbBlue
                    lst.ListSubItems(1).Bold = True
                Else
                    lst.Checked = False
                End If
                RSCUSTOMER.MoveNext
            Wend
        

        

End Sub

Private Sub txtsearch_Change()
 FillSearchGrid (txtsearch.Text)
End Sub



