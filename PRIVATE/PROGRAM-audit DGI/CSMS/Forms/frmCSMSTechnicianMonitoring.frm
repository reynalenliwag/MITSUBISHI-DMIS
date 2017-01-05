VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSTechnicianMonitoring 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Attendance Monitoring"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSTechnicianMonitoring.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdWIT 
      Caption         =   "W/IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6330
      TabIndex        =   19
      ToolTipText     =   "With IT"
      Top             =   6330
      Width           =   735
   End
   Begin VB.CommandButton CmdAOT 
      Caption         =   "AOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   17
      ToolTipText     =   "Authorized OverTime"
      Top             =   6330
      Width           =   705
   End
   Begin VB.CommandButton cmdAbsent 
      Caption         =   "Absent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   15
      ToolTipText     =   "Absent"
      Top             =   6330
      Width           =   735
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "IN/OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   14
      ToolTipText     =   "Show In And Out "
      Top             =   6330
      Width           =   765
   End
   Begin VB.CommandButton CmdOt 
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1380
      TabIndex        =   13
      ToolTipText     =   "Over Time"
      Top             =   6330
      Width           =   705
   End
   Begin VB.CommandButton cmdunderim 
      Caption         =   "UT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2100
      TabIndex        =   12
      ToolTipText     =   "Under Time"
      Top             =   6330
      Width           =   735
   End
   Begin VB.CommandButton cmdAUT 
      Caption         =   "AUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2850
      TabIndex        =   11
      ToolTipText     =   "Authorize Under Time"
      Top             =   6330
      Width           =   705
   End
   Begin VB.CommandButton cmdlate 
      Caption         =   "Late"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3540
      TabIndex        =   10
      ToolTipText     =   "Late"
      Top             =   6330
      Width           =   675
   End
   Begin VB.CommandButton cmdAWL 
      Caption         =   "AWL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   "Absent With Leave"
      Top             =   6330
      Width           =   705
   End
   Begin VB.Frame Frame3 
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   -150
      TabIndex        =   1
      Top             =   8700
      Width           =   7845
      Begin MSComctlLib.ListView listPm 
         Height          =   2805
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4948
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "In PM"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Out Pm"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   750
      Width           =   10965
      Begin MSComctlLib.ListView listAm 
         Height          =   4275
         Left            =   150
         TabIndex        =   35
         Top             =   345
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Today"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "In Am"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Out Am"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "In PM"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Out Pm"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView ListOt 
         Height          =   4275
         Left            =   150
         TabIndex        =   3
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "In Ot"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Out Ot "
            Object.Width           =   2999
         EndProperty
      End
      Begin MSComctlLib.ListView ListAOT 
         Height          =   4275
         Left            =   150
         TabIndex        =   16
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Authorize OT"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListUnderTime 
         Height          =   4275
         Left            =   150
         TabIndex        =   4
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Undertime"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView listLate 
         Height          =   4275
         Left            =   150
         TabIndex        =   6
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Late"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListAWL 
         Height          =   4275
         Left            =   150
         TabIndex        =   7
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "AbsentWleave"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView listWIT 
         Height          =   4275
         Left            =   150
         TabIndex        =   18
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "withIT"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListAll 
         Height          =   4275
         Left            =   150
         TabIndex        =   20
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "In Am "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Out Am"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "In PM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Out PM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "In OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Out OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Undertime"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Authorize UT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Absent W/Leave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Absent "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Authorize OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "With IT"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView listAbsent 
         Height          =   4275
         Left            =   150
         TabIndex        =   8
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Absent"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView listAUT 
         Height          =   4275
         Left            =   150
         TabIndex        =   5
         Top             =   350
         Width           =   10550
         _ExtentX        =   18600
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Authorize UT"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   30
      TabIndex        =   31
      Top             =   -30
      Width           =   10995
      Begin VB.CommandButton CmdAll 
         Caption         =   "All"
         Height          =   435
         Left            =   9540
         TabIndex        =   32
         ToolTipText     =   "Show All Result"
         Top             =   180
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DtpAtend 
         Height          =   375
         Left            =   1380
         TabIndex        =   33
         Top             =   180
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57999361
         CurrentDate     =   39220
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Attendance"
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
         Left            =   90
         TabIndex        =   34
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   30
      TabIndex        =   21
      Top             =   5460
      Width           =   10965
      Begin VB.OptionButton Option8 
         Caption         =   "Authorized OT"
         Height          =   285
         Left            =   9390
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   270
         Width           =   1515
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Absent W/ Leave"
         Height          =   285
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   270
         Width           =   1815
      End
      Begin VB.OptionButton Option9 
         Caption         =   "With/IT"
         Height          =   285
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Absent"
         Height          =   285
         Left            =   5610
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Late"
         Height          =   285
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Authorize UT"
         Height          =   285
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   1395
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Uder Time"
         Height          =   285
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Over Time"
         Height          =   285
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In/Out"
         Height          =   285
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   270
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmCSMSTechnicianMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Thedate                                            As String

Sub fillAm()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,datetoday,inam,outam,inpm,outpm FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listAm.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = listAm.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!EMPNO)
            ITEM.SubItems(2) = Null2String(!Technician)
            ITEM.SubItems(3) = Null2String(!TECH_NAME)
            ITEM.SubItems(4) = Null2String(!datetoday)
            ITEM.SubItems(5) = Null2String(!INAM)
            ITEM.SubItems(6) = Null2String(!OUTAM)
            ITEM.SubItems(7) = Null2String(!InPM)
            ITEM.SubItems(8) = Null2String(!outpm)

            If ITEM.SubItems(5) = "" Then
                ITEM.SubItems(5) = "Absent "
            End If
            If ITEM.SubItems(7) = "" Then
                ITEM.SubItems(7) = "Absent"
            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub FillPM()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT * FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listAm.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = listAm.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!EMPNO)
            ITEM.SubItems(2) = Null2String(!Technician)
            ITEM.SubItems(3) = Null2String(!TECH_NAME)
            ITEM.SubItems(4) = Null2String(!datetoday)
            ITEM.SubItems(5) = Format(!INAM, "hh:mm:ss ampm")
            ITEM.SubItems(6) = Format(!OUTAM, "hh:mm:ss ampm")
            If ITEM.SubItems(5) = "" Then
                ITEM.SubItems(5) = "Absent "
            End If
            If ITEM.SubItems(6) = "" Then
                ITEM.SubItems(6) = "Absent"
            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillOT()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,inot,outot FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListOt.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!inot) <> "" Then
                Set ITEM = ListOt.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = Null2String(!inot)
                ITEM.SubItems(5) = Null2String(!outot)

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillUndertime()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,undertime FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListUnderTime.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!undertime) = "1" Then
                Set ITEM = ListUnderTime.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Undertime"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillAuthorizeUT()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,authorizeut FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listAUT.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!authorizeut) = True Then
                Set ITEM = listAUT.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Authorize UT"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillLate()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,late FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listLate.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!late) = "1" Then
                Set ITEM = listLate.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Late"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillAWL()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,absentWleave FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListAWL.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!absentwleave) = True Then
                Set ITEM = ListAWL.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Absent W/Leave"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillAbsent()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,absent FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listAbsent.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!absent) = "1" Then
                Set ITEM = listAbsent.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Absent"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillAOT()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,authorizeot FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListAOT.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!authorizeot) = True Then
                Set ITEM = ListAOT.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "Unauthorize OT"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub fillWIT()                                         'WIT IT
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT empno,technician,tech_name,withit FROM CSMS_vw_Tech_Attendance WHERE DateToday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listWIT.ListItems.Clear
    cnt = 0
    With RS
        Do While Not .EOF
            cnt = cnt + 1
            If Null2String(!withit) = True Then
                Set ITEM = listWIT.ListItems.Add(, , cnt)
                ITEM.SubItems(1) = Null2String(!EMPNO)
                ITEM.SubItems(2) = Null2String(!Technician)
                ITEM.SubItems(3) = Null2String(!TECH_NAME)
                ITEM.SubItems(4) = "With IT"

            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing

End Sub

Sub fillAll()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim Mdate                                          As String

    Mdate = DtpAtend

    SQL = "SELECT * FROM CSMS_vw_Tech_Attendance WHERE Datetoday='" & Mdate & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListAll.ListItems.Clear

    cnt = 0

    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = ListAll.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!EMPNO)
            ITEM.SubItems(2) = Null2String(!Technician)
            ITEM.SubItems(3) = Null2String(!TECH_NAME)

            ITEM.SubItems(4) = Format(!INAM, "hh:mm:ss ampm")
            ITEM.SubItems(5) = Format(!OUTAM, "hh:mm:ss ampm")
            ITEM.SubItems(6) = Format(!InPM, "hh:mm:ss ampm")
            ITEM.SubItems(7) = Format(!outpm, "hh:mm:ss ampm")
            ITEM.SubItems(8) = Format(!inot, "hh:mm:ss ampm")
            ITEM.SubItems(9) = Format(!outot, "hh:mm:ss ampm")


            ITEM.SubItems(10) = Null2String(!undertime)
            ITEM.SubItems(11) = Null2String(!authorizeut)
            ITEM.SubItems(12) = Null2String(!absentwleave)
            ITEM.SubItems(13) = Null2String(!absent)
            ITEM.SubItems(14) = Null2String(!authorizeot)
            ITEM.SubItems(15) = Null2String(!withit)
            .MoveNext

            If ITEM.SubItems(4) = "" Then
                ITEM.SubItems(4) = "Absent"
            End If

            If ITEM.SubItems(6) = "" Then
                ITEM.SubItems(6) = "Absent"
            End If


            If ITEM.SubItems(10) = "0" Then
                ITEM.SubItems(10) = ""
            Else
                ITEM.SubItems(10) = "Undertime"
            End If

            If ITEM.SubItems(11) = False Then
                ITEM.SubItems(11) = ""
            Else
                ITEM.SubItems(11) = "Authorize UT"
            End If


            If ITEM.SubItems(12) = False Then
                ITEM.SubItems(12) = ""
            Else
                ITEM.SubItems(12) = "Absent W/ Leave "
            End If

            If ITEM.SubItems(13) = "0" Then
                ITEM.SubItems(13) = ""
            Else
                ITEM.SubItems(13) = "absent"
            End If

            If ITEM.SubItems(14) = False Then
                ITEM.SubItems(14) = ""
                ITEM.SubItems(14) = "Authorize OT"
            End If


            If ITEM.SubItems(15) = False Then
                ITEM.SubItems(15) = ""
            Else
                ITEM.SubItems(15) = "With IT"
            End If

        Loop
    End With

    Set RS = Nothing

End Sub

Private Sub cmdAbsent_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = True
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAbsent
End Sub

Private Sub CmdAll_Click()
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    ListAll.Visible = True

    Call fillAll

    LogAudit "I", "TECHNICIAN ATTENDANCE MONITORING"
End Sub

Private Sub CmdAOT_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    ListAOT.Visible = True
    listWIT.Visible = False
    Call fillAOT
End Sub

Private Sub cmdAUT_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = True
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAuthorizeUT
End Sub

Private Sub cmdAWL_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = True
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAWL
End Sub

Private Sub cmdlate_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = True
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillLate
End Sub

Private Sub CmdOt_Click()
    ListAll.Visible = False
    listAm.Visible = False                            'list of all tech with the in and out
    ListOt.Visible = True
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillOT
End Sub

Private Sub cmdShow_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = True
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAm
End Sub

Private Sub Command1_Click()
    'Fillthenull
End Sub

Private Sub cmdunderim_Click()
    ListAll.Visible = False
    listAm.Visible = False
    ListOt.Visible = False
    ListUnderTime.Visible = True
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillUndertime
End Sub

Private Sub cmdWIT_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    listWIT.Visible = True
    Call fillWIT
End Sub

Private Sub DtpAtend_Change()
    Call fillAll
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DtpAtend.Value = Format(Now, "MM/dd/yyyy")
    CmdAll_Click
    ListOt.Visible = False
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Option1_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = True
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAm
End Sub

Private Sub Option2_Click()
    ListAll.Visible = False
    listAm.Visible = False
    ListOt.Visible = True
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillOT
End Sub

Private Sub Option3_Click()
    ListAll.Visible = False
    listAm.Visible = False
    ListOt.Visible = False
    ListUnderTime.Visible = True
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillUndertime
End Sub

Private Sub Option4_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = True
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAuthorizeUT
End Sub

Private Sub Option5_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = True
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillLate
End Sub

Private Sub Option6_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = True
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAWL
End Sub

Private Sub Option7_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = True
    ListAOT.Visible = False
    listWIT.Visible = False
    Call fillAbsent
End Sub

Private Sub Option8_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    ListAOT.Visible = True
    listWIT.Visible = False
    Call fillAOT
End Sub

Private Sub Option9_Click()
    ListAll.Visible = False
    ListOt.Visible = False
    listAm.Visible = False
    ListUnderTime.Visible = False
    listAUT.Visible = False
    listLate.Visible = False
    ListAWL.Visible = False
    listAbsent.Visible = False
    ListAOT.Visible = False
    listWIT.Visible = False
    listWIT.Visible = True
    Call fillWIT
End Sub

