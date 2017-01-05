VERSION 5.00
Begin VB.Form frmListAbsences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Absentees for the Day"
   ClientHeight    =   6105
   ClientLeft      =   2340
   ClientTop       =   1185
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   7440
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   210
      TabIndex        =   6
      Top             =   1380
      Width           =   6945
   End
   Begin VB.Label Label6 
      Caption         =   "Afternoon"
      Height          =   375
      Left            =   5775
      TabIndex        =   5
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Morning"
      Height          =   375
      Left            =   4725
      TabIndex        =   4
      Top             =   1035
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Employee's Name"
      Height          =   375
      Left            =   1155
      TabIndex        =   3
      Top             =   1035
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "EMPNO"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   450
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "LIST OF ABSENTEES FOR THE DAY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   525
      TabIndex        =   0
      Top             =   60
      Width           =   6375
   End
End
Attribute VB_Name = "frmListAbsences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Label2.Caption = Format(Date, "dddddd")
    Dim rsEmpInfos                                     As ADODB.Recordset
    Set rsCard = New ADODB.Recordset
    rsCard.Open "Select * from HRMS_Attend where DateToday = '" & Format(Now, "short date") & "'", GCONDMIS
    If Not rsCard.EOF And Not rsCard.BOF Then
        rsCard.MoveFirst
        Do Until rsCard.EOF
            Set rsEmpInfos = New ADODB.Recordset
            'Set rsEmpInfos = gconDMIS.Execute("Select * from HRMS_EmpInfo where DivCode = " & thedivcode & " and ACTIVEINACTIVE = 'A' and EmpNo = " & N2Str2Null(rsCard!empno))
            Set rsEmpInfos = GCONDMIS.Execute("Select * from HRMS_EmpInfo where ACTIVEINACTIVE = 'A' and EmpNo = " & N2Str2Null(rsCard!empno))
            If Not rsEmpInfos.EOF And Not rsEmpInfos.BOF Then
                'If Null2String(rsEmpInfos!divcode) = thedivcode Then
                A1 = Space(6): P1 = Space(6):
                If Time < #1:00:00 PM# Then
                    If Null2String(rsCard!InAm) = "" Then
                        A1 = "Absent"
                    End If
                Else
                    If Null2String(rsCard!InAm) = "" Then
                        A1 = "Absent"
                    End If

                    If Null2String(rsCard!InPm) = "" Then
                        P1 = "Absent"
                    End If
                End If

                If A1 = "Absent" Or P1 = "Absent" Then
                    SP = 30 - Len(Null2String(rsEmpInfos!LastName) & ", " & Null2String(rsEmpInfos!FirstName))
                    ST = " " & Format(Null2String(rsEmpInfos!empno), "000") & "  " & Null2String(rsEmpInfos!LastName) & ", " & Null2String(rsEmpInfos!FirstName) & Space(SP) & A1 & Space(3) & P1
                    List1.AddItem ST
                End If
                'End If
            End If
            rsCard.MoveNext
        Loop
    End If
End Sub
