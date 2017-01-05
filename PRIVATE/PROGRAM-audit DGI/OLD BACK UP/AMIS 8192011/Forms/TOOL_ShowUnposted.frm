VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTOOL_ShowUnposted 
   Caption         =   "Form2"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6906
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTOOL_ShowUnposted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rsCSMS_REPOR As ADODB.Recordset
Dim rsPMIS_ORD_HD As ADODB.Recordset
Dim rsSMIS_PURCHAGREE As ADODB.Recordset

Set rsCSMS_REPOR = New ADODB.Recordset
Set rsCSMS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR Order by REP_OR ASC")
If Not rsCSMS_REPOR.EOF And Not rsCSMS_REPOR.BOF Then
   rsCSMS_REPOR.MoveFirst
   
End If
End Sub
