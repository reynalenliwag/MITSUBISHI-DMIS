VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      ClipControls    =   0   'False
      Height          =   4275
      Left            =   3090
      TabIndex        =   3
      Top             =   750
      Width           =   5775
      Begin MSComctlLib.ListView ListView1 
         Height          =   3945
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   6959
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   30
      TabIndex        =   2
      Top             =   5100
      Width           =   8865
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   3450
         TabIndex        =   8
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      ClipControls    =   0   'False
      Height          =   4275
      Left            =   30
      TabIndex        =   1
      Top             =   780
      Width           =   2955
      Begin MSComctlLib.ListView ListView2 
         Height          =   3975
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   7011
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8865
      Begin VB.ComboBox cbomoduletype 
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
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   1815
      End
      Begin VB.ComboBox cboModule 
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
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Module Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4020
         TabIndex        =   9
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Module Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsmainmodule                       As ADODB.Recordset
Dim rsmoduletype                       As ADODB.Recordset

Private Sub Combo1_Change()

End Sub

Private Sub initcombo()
    cboModule.Clear
End Sub
Private Sub cleargrid()
    
End Sub

Private Sub cboModule_Change()
    showmoduletype
End Sub


Private Sub cboModule_GotFocus()
    showmoduletype
End Sub

Private Sub Form_load()
    initcombo
    showallmodule
End Sub
Sub showallmodule()
    Dim rsmainmodule As New ADODB.Recordset
    Set rsmainmodule = gconDMIS.Execute("Select distinct mainmodulename from all_rams_modules")
    If Not rsmainmodule.EOF And Not rsmainmodule.BOF Then
        rsmainmodule.MoveFirst: cboModule.Clear
        Do While Not rsmainmodule.EOF
            cboModule.AddItem Null2String(rsmainmodule!mainmodulename)
            rsmainmodule.MoveNext
        Loop
    End If
End Sub
Sub showmoduletype()
    Dim rsmoduletype As New ADODB.Recordset
    Set rsmoduletype = gconDMIS.Execute("select Distinct module_type from all_rams_modules")
    If Not rsmoduletype.EOF And Not rsmoduletype.BOF Then
        rsmoduletype.MoveFirst: cbomoduletype.Clear
        Do While Not rsmoduletype.EOF
            cbomoduletype.AddItem N2String(rsmoduletype!module_type)
            rsmoduletype.MoveNext
        Loop
    End If
End Sub

