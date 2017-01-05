VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPMISAC_DataEntryTagByPartNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Numbers By Part Numbers"
   ClientHeight    =   6300
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   7575
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_DataEntryTagByPartNo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7575
   Begin MSFlexGridLib.MSFlexGrid grdTags 
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483637
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      TextStyleFixed  =   3
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "AC_DataEntryTagByPartNo.frx":030A
   End
End
Attribute VB_Name = "frmPMISAC_DataEntryTagByPartNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTAGS                                                            As ADODB.Recordset

Sub rsRefresh()
    Set rsTAGS = New ADODB.Recordset
    rsTAGS.Open "Select * from tags  order by STOCKNO,val(tag) asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    With grdTags
        .Row = 0
        .FormatString = "Part No                  | TAG NO.      | Status | Remarks                                                 | " & _
                        "Duplicate               "
    End With
End Sub

Sub FillGrid()
    Dim kcnt                                                          As Integer
    kcnt = 0
    If Not rsTAGS.EOF And Not rsTAGS.BOF Then
        Screen.MousePointer = 11
        rsTAGS.MoveFirst
        Do While Not rsTAGS.EOF
            kcnt = kcnt + 1
            grdTags.AddItem Null2String(rsTAGS!STOCKNO) & Chr(9) & _
                            Null2String(rsTAGS!Tag) & Chr(9) & _
                            Null2String(rsTAGS!Status) & Chr(9) & _
                            Null2String(rsTAGS!remarks) & Chr(9) & _
                            Null2String(rsTAGS!Duplicate)
            rsTAGS.MoveNext
        Loop
        If kcnt <> 0 Then grdTags.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cleargrid grdTags
    rsRefresh
    InitGrid
    FillGrid
    Screen.MousePointer = 0
End Sub

