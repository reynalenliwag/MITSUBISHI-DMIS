VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Inquiry_CalendarSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Appointment Calendar"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SalesCalendar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11085
   Begin XtremeReportControl.ReportControl lstAppointment 
      Height          =   6285
      Left            =   90
      TabIndex        =   0
      Top             =   1035
      Width           =   10935
      _Version        =   655364
      _ExtentX        =   19288
      _ExtentY        =   11086
      _StockProps     =   64
      BorderStyle     =   4
      ShowGroupBox    =   -1  'True
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      ShowItemsInGroups=   -1  'True
      ShowFooter      =   -1  'True
   End
   Begin VB.ListBox lstMonth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      IntegralHeight  =   0   'False
      Left            =   6840
      TabIndex        =   2
      Top             =   45
      Width           =   4170
   End
   Begin VB.Label LabelPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   930
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   6630
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_CalendarSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    ReportControlAddColumnHeader lstAppointment, _
                                 "Date, Time, ProspectName, Make, Color, SAE"
    ResizeColumnHeader lstAppointment, "10,15,10,10,10,10,10"
    lstAppointment.PaintManager.TextFont.Size = 9
    lstAppointment.PaintManager.TextFont.Name = "Arial"
    ReportControlPaintManager lstAppointment

    With lstMonth
        .AddItem ("January")
        .AddItem ("February")
        .AddItem ("March")
        .AddItem ("April")
        .AddItem ("May")
        .AddItem ("June")
        .AddItem ("July")
        .AddItem ("August")
        .AddItem ("September")
        .AddItem ("October")
        .AddItem ("November")
        .AddItem ("December")
    End With
    lstMonth.selected(Month(Now) - 1) = True
End Sub


Sub FillAppointments(Where)
    Dim SQL                            As String
    SQL = "SELECT  Convert(varchar, CSA.StartDateTime,101),   Convert(varchar, CSA.StartDateTime ,108)+ ' - '+ Convert(varchar, CSA.EndDateTime ,108) , " & _
        "  CP.AcctName, CSA.Model, CSA.Color, " & _
        "  CSA.SAE,  CP.CUSCDE " & _
        "  FROM        CRIS_SalesAppointments   CSA " & _
        "  INNER JOIN  CRIS_Prospects  CP ON CSA.ProspectID = CP.ProspectID " & Where


    Dim rs                             As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
    flex_FillReportView rs, lstAppointment

End Sub


Private Sub lstMonth_Click()
    FillAppointments (Replace(" WHERE Month(StartDateTime)=@XXX AND YEAR(StartDateTime)= YEAR(GETDATE())", "@XXX", lstMonth.ListIndex + 1))
    LabelPreview.Caption = " Monthly Sales (Appoinment) Calendar For:" & lstMonth.Text
End Sub
