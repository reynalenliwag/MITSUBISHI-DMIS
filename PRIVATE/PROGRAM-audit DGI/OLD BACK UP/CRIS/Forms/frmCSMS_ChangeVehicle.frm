VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMS_ChangeVehicle 
   Caption         =   "Change Vehicle In Repair Order"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "frmCSMS_ChangeVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7365
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   90
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwVehicle 
      Height          =   2235
      Left            =   30
      TabIndex        =   1
      Top             =   540
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PlateNO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Conduction_Sticker"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "VIN"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "NOTE:  Double click the vehicle you want to replace on the existing vehicle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   30
      TabIndex        =   6
      Top             =   3060
      Width           =   7335
   End
   Begin VB.Label lblRONO 
      BackColor       =   &H000000FF&
      Caption         =   "lblRONO"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF0000&
      Height          =   165
      Left            =   30
      TabIndex        =   3
      Top             =   2820
      Width           =   7305
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "Search  Plate No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   7305
   End
End
Attribute VB_Name = "frmCSMS_ChangeVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCUSCDE                                            As String

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    StoreMemVars
End Sub

Sub StoreMemVars()
    Dim rsVehicle                                      As ADODB.Recordset
    Dim rsGetCusCde                                    As ADODB.Recordset
    Dim vLblCapation                                   As String
    Dim ITEM                                           As ListItem

    Set rsGetCusCde = gconDMIS.Execute("Select * from CSMS_RepairOrder where RO_NO = '" & frmCSMSServiceCounter.lblROChange.Caption & "'")
    If Not rsGetCusCde.EOF And Not rsGetCusCde.BOF Then
        vCUSCDE = Null2String(rsGetCusCde!ACCT_NO)
    Else
        MsgBox "Please Select Repair Order..", vbInformation, "INFORMATION"
        Unload Me
    End If

    Me.lvwVehicle.Sorted = True: Me.lvwVehicle.ListItems.Clear: Me.lvwVehicle.Enabled = False
    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select PLATE_NO,MODEL,VCOND_NO,VIN from CSMS_Cusveh where CUSCDE = '" & LTrim(Trim(vCUSCDE)) & "'")

    If Not rsVehicle.EOF And Not rsVehicle.BOF Then
        Do While Not rsVehicle.EOF
            Set ITEM = lvwVehicle.ListItems.Add(, , Null2String(rsVehicle!PLATE_NO))
            ITEM.SubItems(1) = Null2String(rsVehicle!MODEL)
            ITEM.SubItems(2) = Null2String(rsVehicle!VCOND_NO)
            ITEM.SubItems(3) = Null2String(rsVehicle!Vin)
            rsVehicle.MoveNext
        Loop
        Me.lvwVehicle.Enabled = True: Me.lvwVehicle.Sorted = False: Me.lvwVehicle.Refresh
    End If
    Set rsVehicle = Nothing
End Sub

Private Sub lvwVehicle_DblClick()
    Dim vPLATENO                                       As String
    Dim VMODEL                                         As String
    Dim ans                                            As String
    Dim vVin                                           As String


    ans = MsgBox("Are you sure for this vehicle Change?", vbQuestion + vbYesNo)
    If ans = vbYes Then
        vPLATENO = lvwVehicle.SelectedItem.Text
        VMODEL = lvwVehicle.SelectedItem.SubItems(1)
        vVin = lvwVehicle.SelectedItem.SubItems(3)

        SQL_STATEMENT = "Update CSMS_RepairOrder set PLATE_NO = " & N2Str2Null(vPLATENO) & ", MODEl = " & N2Str2Null(VMODEL) & " where RO_NO = " & N2Str2Null(lblRONO) & ""
        gconDMIS.Execute (SQL_STATEMENT)

        SQL_STATEMENT = "Update CSMS_Repor set PLATE_NO = " & N2Str2Null(vPLATENO) & ", MODEL = " & N2Str2Null(VMODEL) & ", VIN = " & N2Str2Null(vVin) & " where Rep_or = " & N2Str2Null(lblRONO) & ""
        gconDMIS.Execute (SQL_STATEMENT)

        ShowSuccessFullyUpdated
        lblRONO.Caption = ""
        Unload Me
    Else
        Exit Sub
    End If
End Sub

Private Sub txtsearch_Change()
    Dim rsSearch                                       As ADODB.Recordset
    Dim vplate_no                                      As String

    If txtSearch = "" Then
        lvwVehicle.Enabled = False
        lvwVehicle.Sorted = False: lvwVehicle.ListItems.Clear
        Set rsSearch = gconDMIS.Execute("select PLATE_NO,MODEL,VCOND_NO from CSMS_Cusveh where CUSCDE = '" & LTrim(Trim(vCUSCDE)) & "'")
        If Not (rsSearch.EOF And rsSearch.BOF) Then
            Listview_Loadval Me.lvwVehicle.ListItems, rsSearch
            lvwVehicle.Refresh
        End If
        lvwVehicle.Enabled = True
    Else
        vplate_no = UCase(txtSearch.Text)

        If vplate_no <> "" Then
            lvwVehicle.Enabled = False
            lvwVehicle.Sorted = False: lvwVehicle.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("select PLATE_NO,MODEL,VCOND_NO from CSMS_Cusveh where CUSCDE = '" & LTrim(Trim(vCUSCDE)) & "' and PLATE_NO like '" & vplate_no & "%'")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwVehicle.ListItems, rsSearch
                lvwVehicle.Refresh
            End If
            lvwVehicle.Enabled = True
        End If
    End If
End Sub
