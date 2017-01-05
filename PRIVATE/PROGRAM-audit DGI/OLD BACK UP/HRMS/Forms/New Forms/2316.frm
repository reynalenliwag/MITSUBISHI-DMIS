VERSION 5.00
Begin VB.Form frmHRMS_2316 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2316 Form"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "2316.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT 2316 FORM"
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
      Left            =   720
      TabIndex        =   3
      Top             =   930
      Width           =   2685
   End
   Begin VB.ComboBox cboName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   555
   End
   Begin VB.Label LABEMPNO 
      Alignment       =   2  'Center
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmHRMS_2316"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    Dim xlApp                                                         As Excel.Application
    Dim xlsheet                                                       As Excel.Worksheet
    Dim xlbook                                                        As Excel.Workbook
    
    If cboName.Text = "" Or cboyear.Text = "" Then
        Exit Sub
    End If
    
    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "2316.XLT")
    Set xlsheet = xlbook.Worksheets(1)
    
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & GetEmployeeNumber(cboName.Text) & "'")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        xlsheet.Shapes("RECTANGLE 17").TextFrame.Characters.Text = " " & Left(cboyear, 1) & "  " & Mid(cboyear, 2, 1) & "  " & Mid(cboyear, 3, 1) & "  " & Right(cboyear, 1)
        xlsheet.Shapes("RECTANGLE 23").TextFrame.Characters.Text = Left(GET_STRING(GET_TIN(Null2String(rsEmpInfo!EMPNO))), 3)
        xlsheet.Shapes("RECTANGLE 25").TextFrame.Characters.Text = Mid(GET_STRING(GET_TIN(Null2String(rsEmpInfo!EMPNO))), 4, 3)
        xlsheet.Shapes("RECTANGLE 27").TextFrame.Characters.Text = Mid(GET_STRING(GET_TIN(Null2String(rsEmpInfo!EMPNO))), 7, 3)
        xlsheet.Shapes("RECTANGLE 130").TextFrame.Characters.Text = GET_ADDRESS(Null2String(rsEmpInfo!EMPNO))
        xlsheet.Shapes("RECTANGLE 131").TextFrame.Characters.Text = Null2String(rsEmpInfo!lastname) + "," + Null2String(rsEmpInfo!FIRSTNAME) + "," + Null2String(rsEmpInfo!MIDDLENAME)
        xlsheet.Shapes("RECTANGLE 18").TextFrame.Characters.Text = GET_ADDRESS(Null2String(rsEmpInfo!EMPNO))
        xlsheet.Shapes("RECTANGLE 248").TextFrame.Characters.Text = "  " & MONTH(Null2String(rsEmpInfo!BIRTHDATE)) & "       " & Day(Null2String(rsEmpInfo!BIRTHDATE)) & "        " & YEAR(Null2String(rsEmpInfo!BIRTHDATE))
        If Left(Null2String(rsEmpInfo!EXSTATUS), 1) = "H" Then
            xlsheet.Shapes("TEXT BOX 351").TextFrame.Characters.Text = " X"
        ElseIf Left(Null2String(rsEmpInfo!EXSTATUS), 1) = "M" Then
            xlsheet.Shapes("TEXT BOX 352").TextFrame.Characters.Text = " X"
        Else
            xlsheet.Shapes("TEXT BOX 349").TextFrame.Characters.Text = " X"
        End If
        xlsheet.Shapes("RECTANGLE 65").TextFrame.Characters.Text = "  " + Mid(GET_STRING(GET_COMPANY_TINNO), 1, 3)
        xlsheet.Shapes("RECTANGLE 67").TextFrame.Characters.Text = "  " + Mid(GET_STRING(GET_COMPANY_TINNO), 4, 3)
        xlsheet.Shapes("RECTANGLE 69").TextFrame.Characters.Text = "  " + Mid(GET_STRING(GET_COMPANY_TINNO), 7, 3)
        xlsheet.Shapes("RECTANGLE 81").TextFrame.Characters.Text = GET_COMPANY_NAME()
        xlsheet.Shapes("RECTANGLE 82").TextFrame.Characters.Text = GET_COMPANY_ADDRESS()
        xlsheet.Shapes("RECTANGLE 212").TextFrame.Characters.Text = "X"
    Else
        Exit Sub
    End If
    
    Set rsEmpInfo = Nothing
    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    DrawXPCtl Me
    'FillcboYear cboyear
    fillcombo_up cboyear
    
    Call FillNames
End Sub

Sub FillNames()
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME FROM HRMS_EMPINFO ORDER BY LASTNAME + ', ' + FIRSTNAME ASC")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.EOF Then
        Combo_Loadval cboName, rsEmpInfo
    End If
    Set rsEmpInfo = Nothing
End Sub

Function GetEmployeeNumber(NAME As String) As String
    GetEmployeeNumber = ""
    
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE LASTNAME + ', ' + FIRSTNAME = '" & NAME & "'")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        GetEmployeeNumber = Null2String(rsEmpInfo!EMPNO)
    End If
    Set rsEmpInfo = Nothing
End Function

Function GET_STRING(ACCOUNTNO As String) As String
    Dim X                                                             As Integer
    Dim AMOUNTSTRING                                                  As String
    AMOUNTSTRING = ""
    For X = 1 To Len(ACCOUNTNO)
        If IsNumeric(Mid(ACCOUNTNO, X, 1)) Then
            AMOUNTSTRING = AMOUNTSTRING & Mid(ACCOUNTNO, X, 1)
        End If
    Next
    GET_STRING = CStr(AMOUNTSTRING)
End Function

Function GET_TIN(EMPNO As String) As String
    GET_TIN = ""
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT TINNO FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_TIN = Null2String(rsTemp!tinno)
    End If
    GET_TIN = RTrim(LTrim(GET_TIN))
    Set rsTemp = Nothing
End Function

Function GET_ADDRESS(EMPNO As String) As String
    GET_ADDRESS = ""
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT ADDRESS FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_ADDRESS = Null2String(rsTemp!ADDRESS)
    End If
    GET_ADDRESS = RTrim(LTrim(GET_ADDRESS))
    Set rsTemp = Nothing
End Function

Function GET_COMPANY_TINNO() As String
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT COMPANYTINNO FROM ALL_PROFILE WHERE MODULENAME = 'HRMS'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_COMPANY_TINNO = Null2String(rsTemp!companytinno)
    End If
    Set rsTemp = Nothing
End Function

Function GET_COMPANY_NAME() As String
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT COMPANYNAME FROM ALL_PROFILE WHERE MODULENAME = 'HRMS'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_COMPANY_NAME = Null2String(rsTemp!CompanyName)
    End If
    Set rsTemp = Nothing
End Function

Function GET_COMPANY_ADDRESS() As String
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT COMPANYADDRESS FROM ALL_PROFILE WHERE MODULENAME = 'HRMS'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_COMPANY_ADDRESS = Null2String(rsTemp!Companyaddress)
    End If
    Set rsTemp = Nothing
End Function
