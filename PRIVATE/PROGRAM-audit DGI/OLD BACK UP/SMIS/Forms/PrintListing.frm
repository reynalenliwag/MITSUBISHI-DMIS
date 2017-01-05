VERSION 5.00
Begin VB.Form frmSMIS_ReportChoice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   Icon            =   "PrintListing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "This Customer List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   780
      TabIndex        =   1
      Top             =   810
      Width           =   3165
   End
   Begin VB.OptionButton Option1 
      Caption         =   "All Customer List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   780
      TabIndex        =   0
      Top             =   480
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "Select Option From The List For Printing Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   3945
   End
End
Attribute VB_Name = "frmSMIS_ReportChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public REPORTNAME                      As String


Private Sub Form_Load()
    Select Case REPORTNAME


    End Select

End Sub
