VERSION 5.00
Begin VB.Form frmCSMSAppointmentDiary 
   Caption         =   "Appointment Diary Report"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2280
      MouseIcon       =   "frmAppointmentDiary.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmAppointmentDiary.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1020
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1320
      MouseIcon       =   "frmAppointmentDiary.frx":059D
      MousePointer    =   99  'Custom
      Picture         =   "frmAppointmentDiary.frx":06EF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1020
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Appointment Diary:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   1950
   End
End
Attribute VB_Name = "frmCSMSAppointmentDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
