VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSYrMkMlEgn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9525
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   6390
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5175
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
      Left            =   5640
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5175
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   4875
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5175
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
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5175
      Width           =   705
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
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
      Left            =   7155
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5175
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
      Left            =   8670
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5175
      Width           =   705
   End
   Begin VB.PictureBox Picture2 
      Height          =   1380
      Left            =   240
      ScaleHeight     =   1320
      ScaleWidth      =   4365
      TabIndex        =   14
      Top             =   1860
      Visible         =   0   'False
      Width           =   4425
      Begin VB.CommandButton cmdCancel1 
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
         Height          =   645
         Left            =   3525
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   600
         Width           =   630
      End
      Begin VB.CommandButton cmdOk1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2850
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   600
         Width           =   630
      End
      Begin VB.TextBox txt1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1230
         TabIndex        =   16
         Top             =   180
         Width           =   2925
      End
      Begin VB.Label cap1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3735
      Left            =   30
      ScaleHeight     =   3675
      ScaleWidth      =   5205
      TabIndex        =   17
      Top             =   630
      Visible         =   0   'False
      Width           =   5265
      Begin VB.CommandButton cmdCancel2 
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
         Height          =   645
         Left            =   4350
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2925
         Width           =   630
      End
      Begin VB.CommandButton cmdOk2 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3675
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2925
         Width           =   630
      End
      Begin VB.TextBox txtEngineVIN 
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2460
         Width           =   1245
      End
      Begin VB.ComboBox cboAspiration 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2070
         Width           =   1515
      End
      Begin VB.TextBox txtFuelType 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtDisplacement 
         Height          =   315
         Left            =   2040
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1320
         Width           =   1245
      End
      Begin VB.TextBox txtCubic 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txtLiters 
         Height          =   315
         Left            =   2040
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtEnginetype 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine VIN"
         Height          =   285
         Left            =   1110
         TabIndex        =   30
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Aspiration"
         Height          =   285
         Left            =   1260
         TabIndex        =   28
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Type"
         Height          =   285
         Left            =   1230
         TabIndex        =   26
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Inch Displacement"
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   1380
         Width           =   1875
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Centimeters"
         Height          =   285
         Left            =   660
         TabIndex        =   22
         Top             =   1020
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Liters"
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Type"
         Height          =   285
         Left            =   1050
         TabIndex        =   18
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   5520
      ScaleHeight     =   4635
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   210
      Width           =   3795
      Begin VB.Label labengine 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1050
         TabIndex        =   13
         Top             =   3570
         Width           =   2655
      End
      Begin VB.Label labmodel 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1050
         TabIndex        =   12
         Top             =   2430
         Width           =   2655
      End
      Begin VB.Label labmake 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1050
         TabIndex        =   11
         Top             =   1470
         Width           =   2655
      End
      Begin VB.Label labyear 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1050
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "ENGINE :"
         Height          =   315
         Left            =   300
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "MODEL :"
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "MAKE :"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "YEAR :"
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstCategory 
      Height          =   4755
      Left            =   2280
      TabIndex        =   4
      Top             =   180
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8387
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Year"
         Object.Width           =   4762
      EndProperty
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   300
      X2              =   9390
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   2160
      X2              =   2160
      Y1              =   30
      Y2              =   4980
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   1710
      Y2              =   150
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   1860
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1860
      X2              =   1860
      Y1              =   150
      Y2              =   1710
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   210
      Top             =   3510
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   210
      Top             =   3030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   210
      Top             =   2550
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   345
      Left            =   210
      Top             =   2070
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "ENGINE"
      Height          =   285
      Left            =   420
      TabIndex        =   3
      Top             =   3600
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "MODEL"
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "MAKE"
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "YEAR"
      Height          =   225
      Left            =   420
      TabIndex        =   0
      Top             =   2160
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   150
      Stretch         =   -1  'True
      Top             =   150
      Width           =   1725
   End
End
Attribute VB_Name = "FrmCSMSYrMkMlEgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoad As ADODB.Recordset
Dim xENGINE As String
Dim xLiters As String
Dim xCubic As String
Dim xDisplacement As String
Dim xFuelType As String
Dim xAspiration As String
Dim xEngineVIN As String

Private Sub cmdAdd_Click()
    If Shape1.Visible = True Then
        cap1.Caption = "Year ": txt1 = ""
        Picture2.Visible = True
        txt1.SetFocus
        Picture3.Visible = False
    ElseIf Shape2.Visible = True Then
        cap1.Caption = "Make ": txt1 = ""
        Picture2.Visible = True
        txt1.SetFocus
        Picture3.Visible = False
    ElseIf Shape3.Visible = True Then
        cap1.Caption = "Model ": txt1 = ""
        Picture2.Visible = True
        txt1.SetFocus
        Picture3.Visible = False
    ElseIf Shape4.Visible = True Then
        InitEngine
        Picture2.Visible = False
        Picture3.Visible = True
    End If
End Sub

Private Sub cmdBack_Click()
    If Shape4.Visible = True Then
        Shape4.Visible = False
        Shape3.Visible = True
        cmdNext.Caption = "Next"
        lstCategory.ColumnHeaders(1).Text = "Model"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Model from CSMIOS_S_MODEL order by model asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape3.Visible = True Then
        Shape3.Visible = False
        Shape2.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Make"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Make from ALL_make order by make asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape2.Visible = True Then
        Shape2.Visible = False
        Shape1.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Year"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select yeer from ALL_year order by yeer asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel1_Click()
    Picture2.Visible = False
End Sub

Private Sub cmdNext_Click()
    If Shape1.Visible = True Then
        Shape1.Visible = False
        Shape2.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Make"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Make from All_Make order by make asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape2.Visible = True Then
        Shape2.Visible = False
        Shape3.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Model"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Model from CSMIOS_S_MODEL order by model asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape3.Visible = True Then
        Shape3.Visible = False
        Shape4.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Engine"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Engine from ALL_Engine order by engine asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
        cmdNext.Caption = "Finish"
    ElseIf Shape4.Visible = True Then
        With FrmCSMSAddVehicle
            .txtyear = labyear.Caption
            .txtMake = labmake.Caption
            .txtModel = labmodel.Caption
            .txtEngine = labengine.Caption
            .labVechicle.Caption = Trim(labyear.Caption) & "  " & Trim(labmake.Caption) & "  " & Trim(labmodel.Caption) & "  " & Trim(labengine.Caption)
        End With
        Unload Me
    End If
    'If cmdNext.Caption = "Finish" Then
    'End If
End Sub


Private Sub cmdOk1_Click()
    If Shape1.Visible = True Then
        With Me.lstCategory
            .ListItems.Add 1, , txt1.Text
            .ListItems(1).Selected = True
        End With
        labyear.Caption = txt1.Text
        cmdCancel1.Value = True
        If txt1.Text <> "" Then
            gconDMIS.Execute "Insert into All_Year " & _
                           " (yeer)" & _
                           " values('" & txt1.Text & "')"
        End If
    ElseIf Shape2.Visible = True Then
        With Me.lstCategory
            .ListItems.Add 1, , txt1.Text
            .ListItems(1).Selected = True
        End With
        labmake.Caption = txt1.Text
        cmdCancel1.Value = True
        If txt1.Text <> "" Then
            gconDMIS.Execute "Insert into All_Make " & _
                           " (make)" & _
                           " values('" & txt1.Text & "')"
        End If
    ElseIf Shape3.Visible = True Then
        With Me.lstCategory
            .ListItems.Add 1, , txt1.Text
            .ListItems(1).Selected = True
        End With
        labmodel.Caption = txt1.Text
        cmdCancel1.Value = True
        If txt1.Text <> "" Then
            gconDMIS.Execute "Insert into CSMIOS_S_MODEL " & _
                           " (model)" & _
                           " values('" & txt1.Text & "')"
        End If
    End If
End Sub

Private Sub cmdCancel2_Click()
    Picture3.Visible = False
End Sub


Private Sub cmdOk2_Click()
    With Me.lstCategory
        .ListItems.Add 1, , txtEnginetype.Text
        .ListItems(1).Selected = True
    End With
    labengine.Caption = txtEnginetype.Text
    cmdCancel2.Value = True
    If txtEnginetype.Text <> "" Then
        xENGINE = N2Str2Null(txtEnginetype)
        xLiters = N2Str2Null(txtLiters)
        xCubic = N2Str2Null(txtCubic)
        xDisplacement = N2Str2Null(txtDisplacement)
        xFuelType = N2Str2Null(txtFuelType)
        xAspiration = N2Str2Null(cboAspiration.Text)
        xEngineVIN = N2Str2Null(txtEngineVIN)
        gconDMIS.Execute "Insert into All_Engine " & _
                       " (ENGINE,Liters,Cubic,Displacement,FuelType,Aspiration,EngineVIN)" & _
                       " values(" & xENGINE & "," & xLiters & "," & xCubic & "," & xDisplacement & "," & xFuelType & "," & xAspiration & "," & xEngineVIN & ")"
    End If
End Sub

Private Sub Form_Load()
    Picture2.Left = 3450
    Picture2.Top = 1710
    Picture3.Left = 3390
    Picture3.Top = 690
    InitField
    InitEngine
    Set rsLoad = New ADODB.Recordset
    lstCategory.Sorted = False: lstCategory.ListItems.Clear
    Set rsLoad = gconDMIS.Execute("Select YEER from All_Year order by YEER asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Listview_Loadval Me.lstCategory.ListItems, rsLoad
    End If
End Sub
Sub InitField()
    labyear.Caption = "": labmake.Caption = "": labmodel.Caption = "": labengine.Caption = ""
End Sub
Sub InitEngine()
    txtEnginetype = "": txtLiters = "": txtCubic = "": txtDisplacement = ""
    txtFuelType = "": cboAspiration.ListIndex = -1: txtEngineVIN = ""
End Sub
Private Sub lstCategory_DblClick()
    If Shape1.Visible = True Then
        labyear.Caption = lstCategory.SelectedItem
        Shape1.Visible = False
        Shape2.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Make"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Make from All_Make order by make asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape2.Visible = True Then
        labmake.Caption = lstCategory.SelectedItem
        Shape2.Visible = False
        Shape3.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Model"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Model from CSMIOS_S_MODEL order by model asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        lstCategory.SetFocus
    ElseIf Shape3.Visible = True Then
        labmodel.Caption = lstCategory.SelectedItem
        Shape3.Visible = False
        Shape4.Visible = True
        lstCategory.ColumnHeaders(1).Text = "Engine"
        Set rsLoad = New ADODB.Recordset
        lstCategory.Sorted = False: lstCategory.ListItems.Clear
        Set rsLoad = gconDMIS.Execute("Select Engine from ALL_Engine order by engine asc")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            Listview_Loadval Me.lstCategory.ListItems, rsLoad
        End If
        cmdNext.Caption = "Finish"
        lstCategory.SetFocus
    ElseIf Shape4.Visible = True Then
        'labengine.Caption = lstCategory.SelectedItem
    End If

End Sub
