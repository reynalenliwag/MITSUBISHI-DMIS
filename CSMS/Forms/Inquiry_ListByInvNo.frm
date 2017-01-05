VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCSMS_Inquiry_ListByInvNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INQUIRY BY INVOICE NUMBER"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Inquiry_ListByInvNo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   12375
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   12375
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton Command1 
         Caption         =   "Search RO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   11370
         MouseIcon       =   "Inquiry_ListByInvNo.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Inquiry_ListByInvNo.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Search R.O."
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   6
         Left            =   9060
         TabIndex        =   4
         Tag             =   "DTE_REL"
         Text            =   "Combo1"
         Top             =   90
         Width           =   1845
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Released:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7650
         TabIndex        =   8
         Top             =   150
         Width           =   3465
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   1
         Left            =   1530
         TabIndex        =   5
         Tag             =   "NIYM"
         Text            =   "Combo1"
         Top             =   480
         Width           =   1965
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   525
         Width           =   3645
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   5
         Left            =   5070
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "RECD_BY"
         Top             =   855
         Width           =   2025
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   7
         Left            =   9060
         TabIndex        =   11
         Tag             =   "DTE_COMP"
         Text            =   "Combo1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   2
         Left            =   1530
         TabIndex        =   13
         Tag             =   "PLATE_NO"
         Text            =   "Combo1"
         Top             =   855
         Width           =   1965
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   4
         Left            =   5070
         TabIndex        =   9
         Tag             =   "INVOICE"
         Text            =   "Combo1"
         Top             =   480
         Width           =   2025
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   3
         Left            =   5070
         TabIndex        =   3
         Tag             =   "MODEL"
         Text            =   "Combo1"
         Top             =   90
         Width           =   2025
      End
      Begin VB.ComboBox cbo 
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
         Height          =   345
         Index           =   0
         Left            =   1530
         TabIndex        =   1
         Tag             =   "REP_OR"
         Text            =   "Combo1"
         Top             =   90
         Width           =   1965
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "                    RO #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Width           =   3615
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4470
         TabIndex        =   7
         Top             =   135
         Width           =   2865
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "  INV#:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4470
         TabIndex        =   10
         Top             =   525
         Width           =   2865
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Plate#:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   870
         TabIndex        =   14
         Top             =   900
         Width           =   2835
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Completed:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   7620
         TabIndex        =   12
         Top             =   510
         Width           =   3525
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "      SA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4470
         TabIndex        =   16
         Top             =   900
         Width           =   2865
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdQUERY2 
      Height          =   2055
      Left            =   60
      TabIndex        =   18
      Top             =   4620
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   9
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      FormatString    =   $"Inquiry_ListByInvNo.frx":0797
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdQUERY 
      Height          =   3225
      Left            =   60
      TabIndex        =   17
      ToolTipText     =   "S"
      Top             =   1320
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5689
      _Version        =   393216
      Cols            =   25
      FixedCols       =   3
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      FocusRect       =   2
      SelectionMode   =   1
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   $"Inquiry_ListByInvNo.frx":0826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Inquiry_ListByInvNo.frx":09A1
   End
   Begin VB.Label labTotalRO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9780
      TabIndex        =   20
      Top             =   6210
      Width           =   1455
   End
End
Attribute VB_Name = "frmCSMS_Inquiry_ListByInvNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As ADODB.Recordset
Dim rsRO_DET                                           As ADODB.Recordset
Dim rsEmpNo                                            As ADODB.Recordset
Dim kcnt                                               As Integer

Function GetSACode(XXX As String) As String
    Dim temprs                                         As ADODB.Recordset
    XXX = ReplaceQuote(XXX)


    Set temprs = gconDMIS.Execute("select CODE from CSMS_vw_EmpNo where upper(NAYM)='" & UCase(XXX) & "'")

    If Not (temprs.EOF Or temprs.BOF) Then
        GetSACode = Null2String(temprs!code)
    End If
End Function

Function InvoiceBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                         As Boolean
    Dim rsBClone                                       As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsREPOR.Clone

        rsBClone.Find "Invoice = '" & UCase(str2find) & "'"
        result = Not rsBClone.EOF
        If result Then
            rsREPOR.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    InvoiceBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Function PlateNoBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                         As Boolean
    Dim rsBClone                                       As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsREPOR.Clone

        rsBClone.Find "Plate_No = '" & UCase(str2find) & "'"
        result = Not rsBClone.EOF
        If result Then
            rsREPOR.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    PlateNoBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Function ReporBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                         As Boolean
    Dim rsBClone                                       As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsREPOR.Clone

        rsBClone.Find "rep_or = '" & UCase(str2find) & "'"
        result = Not rsBClone.EOF
        If result Then
            rsREPOR.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    ReporBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Sub FillReporGrid()
    On Error GoTo ErrorCode
    kcnt = 0
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        Screen.MousePointer = 11
        rsREPOR.MoveFirst
        Do While Not rsREPOR.EOF
            kcnt = kcnt + 1
            grdQUERY.AddItem Null2String(rsREPOR!rep_OR) & Chr(9) & _
                             Null2String(rsREPOR!PLATE_NO) & Chr(9) & _
                             Null2String(rsREPOR!invoice) & Chr(9) & _
                             Null2String(rsREPOR!NIYM) & Chr(9) & _
                             Null2String(rsREPOR!MODEL) & Chr(9) & _
                             Null2String(rsREPOR!TERM) & Chr(9) & _
                             Null2String(rsREPOR!recd_by) & Chr(9) & _
                             Null2String(rsREPOR!km_rdg) & Chr(9) & _
                             Null2String(rsREPOR!dte_recd) & Chr(9) & _
                             Null2String(rsREPOR!certific8) & Chr(9) & _
                             N2Str2Zero(rsREPOR!amount) & Chr(9) & _
                             N2Str2Zero(rsREPOR!insamt) & Chr(9) & _
                             N2Str2Zero(rsREPOR!rovat) & Chr(9) & _
                             N2Str2Zero(rsREPOR!labor) & Chr(9) & _
                             N2Str2Zero(rsREPOR!parts) & Chr(9) & _
                             N2Str2Zero(rsREPOR!material) & Chr(9) & _
                             Null2String(rsREPOR!prin_dte) & Chr(9) & _
                             Null2String(rsREPOR!dte_comp) & Chr(9) & _
                             Null2String(rsREPOR!dte_rel) & Chr(9) & _
                             N2Str2Zero(rsREPOR!invbal) & Chr(9) & _
                             Null2String(rsREPOR!Status)

            rsREPOR.MoveNext
            DoEvents
        Loop
        If kcnt <> 0 Then grdQUERY.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cleargrid grdQUERY
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub InitCombo()
    Dim temprs                                         As ADODB.Recordset
    'ro Number
    Set temprs = gconDMIS.Execute("Select REP_OR from CSMS_RepOr where Transtype='R'")
    Combo_Loadval cbo(0), temprs
    '    Customer Name
    Set temprs = gconDMIS.Execute("Select DISTINCT NIYM from CSMS_RepOr where Transtype='R'")
    Combo_Loadval cbo(1), temprs
    'PlateNumber
    Set temprs = gconDMIS.Execute("Select DISTINCT PLATE_NO from CSMS_RepOr where Transtype='R' and PLATE_NO is not null")
    Combo_Loadval cbo(2), temprs
    'Model
    Set temprs = gconDMIS.Execute("Select DISTINCT MODEL from CSMS_RepOr where Transtype='R' and MODEL is not null")
    Combo_Loadval cbo(3), temprs
    'Invoice Number
    Set temprs = gconDMIS.Execute("Select DISTINCT INVOICE from CSMS_RepOr where Transtype='R' and INVOICE is not null")
    Combo_Loadval cbo(4), temprs
    'SA
    Set temprs = gconDMIS.Execute("Select NAYM from CSMS_vw_EmpNo order by lastname asc ")
    Combo_Loadval cbo(5), temprs
    'Date Released
    Set temprs = gconDMIS.Execute("Select  DISTINCT DTE_REL from CSMS_RepOr where Transtype='R' and DTE_REL is not null")
    Combo_Loadval cbo(6), temprs
    'Date Completed
    Set temprs = gconDMIS.Execute("Select DISTINCT DTE_COMP from CSMS_RepOr where Transtype='R' and DTE_COMP is not null")
    Combo_Loadval cbo(7), temprs

End Sub

Sub ShadeControl(oBx As Object, ISTrue As Boolean, Optional ByVal xVal As Variant = vbNullString)
    If ISTrue Then
        oBx.Enabled = True
        oBx.BackColor = vbWhite
    Else
        oBx.Enabled = False
        oBx.BackColor = vbButtonFace
    End If
    If xVal <> vbNullString Then: oBx.Text = xVal
End Sub

Sub VcmdInvoiceSearch()
    On Error GoTo ErrorCode
    Dim FOUNDTEXT                                      As String
    Dim FOUNDNUM                                       As Long
    Dim findStr                                        As String
    FOUNDNUM = 0
    grdQUERY.Col = 1
    MsgSpeechBox "Please Input Invoice Number to Search"
    findStr = InputBoxXP("Please Input Invoice Number to Search", "Find", grdQUERY.Text)
    If findStr <> "" Then
        findStr = Format(findStr, "000000")
        If Not InvoiceBFound(findStr) Then
            MsgSpeechBox "Invoice Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsREPOR.AbsolutePosition
            grdQUERY.Row = FOUNDNUM
            grdQUERY.RowSel = FOUNDNUM
            grdQUERY.SetFocus
            FOUNDTEXT = grdQUERY.Text
            grdQUERY.TopRow = FOUNDNUM
        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub VcmdPlateNSearch()
    On Error GoTo ErrorCode
    Dim FOUNDTEXT                                      As String
    Dim FOUNDNUM                                       As Long
    Dim findStr                                        As String
    FOUNDNUM = 0
    grdQUERY.Col = 2
    MsgSpeechBox "Please Input Plate Number to Search"
    findStr = InputBoxXP("Please Input Plate Number to Search", "Find", grdQUERY.Text)
    If findStr <> "" Then
        If Not PlateNoBFound(findStr) Then
            MsgSpeechBox "Plate Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsREPOR.AbsolutePosition
            grdQUERY.Row = FOUNDNUM
            grdQUERY.RowSel = FOUNDNUM
            grdQUERY.SetFocus
            FOUNDTEXT = grdQUERY.Text
            grdQUERY.TopRow = FOUNDNUM
        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub VcmdRODetails()
    Dim fild                                           As String
    Dim RONumber                                       As String
    Dim YzaCnt                                         As Integer
    YzaCnt = 0
    grdQUERY.Col = 0
    RONumber = grdQUERY.Text
    fild = grdQUERY.Text
    grdQUERY2.ZOrder 0
    cleargrid grdQUERY2

    If RONumber <> "" Then
        Set rsRO_DET = New ADODB.Recordset
        rsRO_DET.Open "select * from CSMS_RO_Det where TRANSTYPE='R' and rep_or = '" & RONumber & "'", gconDMIS
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            rsRO_DET.MoveFirst
            Do While Not rsRO_DET.EOF
                grdQUERY2.AddItem Null2String(rsRO_DET!DETDSC) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!detvol) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!DetPrc) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!DETAMT) & Chr(9) & _
                                  Null2String(rsRO_DET!wCode) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!taxrate) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!discrate) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!TAXVAL) & Chr(9) & _
                                  N2Str2Zero(rsRO_DET!disval) & Chr(9)

                YzaCnt = YzaCnt + 1
                rsRO_DET.MoveNext
            Loop
        End If
        If YzaCnt > 0 Then grdQUERY2.RemoveItem 1
    Else
        Exit Sub
    End If
End Sub

Sub VcmdROSearch()
    On Error GoTo ErrorCode
    Dim FOUNDTEXT                                      As String
    Dim FOUNDNUM, k                                    As Long
    Dim infindstr, findStr, findstr2, findstr3         As String
    FOUNDNUM = 0
    grdQUERY.Col = 0
    MsgSpeechBox "Please Input Repair Order Number to Search"
    infindstr = InputBoxXP("Please Input Repair Order Number to Search", "Find", grdQUERY.Text)
    If infindstr <> "" Then
        If IsNumeric(infindstr) = True Then
            findStr = Format(Left(infindstr, 2), "A-") & Format(Right(infindstr, 6), "000000")
        Else
            For k = 1 To Len(infindstr)
                findstr2 = Mid(infindstr, k, 1)
                If IsNumeric(findstr2) = True Then
                    findstr3 = findstr3 + findstr2
                End If
            Next
            findstr3 = Format(findstr3, "000000")
            findStr = Format(Left(findstr3, 2), "A-") & Format(Right(findstr3, 6), "000000")
        End If
        If Not ReporBFound(findStr) Then
            MsgSpeechBox "Repair Order Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsREPOR.AbsolutePosition
            grdQUERY.Row = FOUNDNUM
            grdQUERY.RowSel = FOUNDNUM
            grdQUERY.SetFocus
            FOUNDTEXT = grdQUERY.Text
            grdQUERY.TopRow = FOUNDNUM
        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub VcmdSearchSA()
    On Error GoTo ErrorCode
    Dim FOUNDTEXT                                      As String
    Dim FOUNDNUM, k                                    As Long
    Dim infindstr, findStr, findstr2, findstr3         As String
    FOUNDNUM = 0
    grdQUERY.Col = 1
    MsgSpeechBox "Please Input Repair Order Number to Search"
    infindstr = InputBoxXP("Please Input Repair Order Number to Search", "Find", grdQUERY.Text)
    If infindstr <> "" Then
        If IsNumeric(infindstr) = True Then
            findStr = Format(Left(infindstr, 2), "A-") & Format(Right(infindstr, 6), "000000")
        Else
            For k = 1 To Len(infindstr)
                findstr2 = Mid(infindstr, k, 1)
                If IsNumeric(findstr2) = True Then
                    findstr3 = findstr3 + findstr2
                End If
            Next
            findstr3 = Format(findstr3, "000000")
            findStr = Format(Left(findstr3, 2), "A-") & Format(Right(findstr3, 6), "000000")
        End If
        If Not ReporBFound(findStr) Then
            MsgSpeechBox "Repair Order Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsREPOR.AbsolutePosition
            grdQUERY.Row = FOUNDNUM
            grdQUERY.RowSel = FOUNDNUM
            grdQUERY.SetFocus
            FOUNDTEXT = grdQUERY.Text
            grdQUERY.TopRow = FOUNDNUM
        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub chk_Click(Index As Integer)
    If chk(Index).Value = 1 Then
        Call ShadeControl(cbo(Index), True)
        If cbo(Index).ListCount > 0 Then cbo(Index).ListIndex = 0
    Else
        Call ShadeControl(cbo(Index), False)
        cbo(Index).ListIndex = -1

    End If
End Sub

Private Sub cmdInvoiceSearch_Click()
    VcmdInvoiceSearch
End Sub

Private Sub cmdPARTSINQUIRYExit_Click()
    Unload Me
End Sub

Private Sub cmdPlateNSearch_Click()
    VcmdPlateNSearch
End Sub

Private Sub cmdROSearch_Click()
    VcmdROSearch
End Sub

Private Sub Command1_Click()
    Dim I                                              As Long
    Dim SearchString1                                  As String
    Dim SACODE                                         As String

    For I = 0 To chk.Count - 1

        If chk(I).Value = 1 Then
            If cbo(I) = "<ALL>" Or cbo(I) <> "" Then
                If I = 5 Then
                    'get sa code
                    SACODE = GetSACode(cbo(5).Text)
                    If Not SACODE = "" Then: SearchString1 = SearchString1 & " UPPER(" & cbo(I).Tag & ")='" & UCase(cbo(I).Text) & "' AND "
                Else
                    SearchString1 = SearchString1 & cbo(I).Tag & "='" & cbo(I).Text & "' AND "
                End If

            End If

        End If
    Next

    If Len(SearchString1) > 0 Then
        SearchString1 = Left(SearchString1, Len(SearchString1) - 4)
        SearchString1 = " AND  " & SearchString1
    End If




    cleargrid grdQUERY
    Me.Caption = "INQUIRY BY INVOICE NUMBER"
    Set rsREPOR = New ADODB.Recordset
    rsREPOR.Open "select * from CSMS_RepOr where Transtype='R' AND Invoice is not null " & SearchString1 & " order by invoice asc", gconDMIS
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        FillReporGrid
    End If
    If grdQUERY.Cols > 1 Then
        VcmdRODetails
        grdQUERY.Col = 1
        grdQUERY.SetFocus
        grdQUERY.ColSel = grdQUERY.Cols - 1

    End If

    LogAudit "I", "INQUIRY BY INVOICE NUMBER"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            grdQUERY2.ZOrder 1
        Case vbKeyF2
            If SearchBy = "RO" Then VcmdROSearch
            If SearchBy = "INV" Then VcmdInvoiceSearch
            If SearchBy = "PLN" Then VcmdPlateNSearch
        Case vbKeyF3
            VcmdRODetails
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitCombo
    cleargrid grdQUERY
    Set rsREPOR = New ADODB.Recordset
    rsREPOR.Open "select * from CSMS_RepOr where Transtype='R' And Invoice is not null order by invoice asc", gconDMIS
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        FillReporGrid
    End If
    If grdQUERY.Cols > 1 Then
        VcmdRODetails
        grdQUERY.Col = 1
        On Error Resume Next
        grdQUERY.SetFocus
        grdQUERY.ColSel = grdQUERY.Cols - 1
    End If
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMS_Inquiry_ListByInvNo = Nothing
End Sub

Private Sub grdQUERY_Click()
    On Error Resume Next
    VcmdRODetails
    grdQUERY.SetFocus
    grdQUERY.ColSel = grdQUERY.Cols - 1
End Sub

