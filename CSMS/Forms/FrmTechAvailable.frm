VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTechAvailable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Status"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lblTech 
      Height          =   4125
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4275
      _ExtentX        =   7541
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmTechAvailable.frx":0000
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " Technician"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "FrmTechAvailable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        Dim rsUpload As ADODB.Recordset
        lblTech.Sorted = False: lblTech.ListItems.Clear
        Set rsUpload = New ADODB.Recordset
        Set rsUpload = gconDMIS.Execute("Select Technician,Tech_Name,IS_TECH_STATUS from [CSMIOS_vw_Technician] Order by [status] Asc")
        If Not rsUpload.EOF And Not rsUpload.BOF Then
           Listview_Loadval Me.lblTech.ListItems, rsUpload
        End If
End Sub
