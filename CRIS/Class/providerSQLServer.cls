VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "providerSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m_pCalendar                   As CalendarControl
Attribute m_pCalendar.VB_VarHelpID = -1
Dim m_Connection                             As New ADODB.Connection
Dim m_bOpened                                As Boolean
Dim m_pHelper                                As New SQLDataHelper
Dim ApType                                   As Integer

Private Sub Class_Initialize()
    m_bOpened = False
End Sub

Public Sub SetCalendar(ByVal pCalendar As CalendarControl, AppointmentType As Integer)
    Set m_pCalendar = pCalendar
    ApType = AppointmentType
    m_pHelper.SetCalendar pCalendar
End Sub
Public Sub OpenDB(ByVal strConnection As String)
    '  On Error GoTo err1
    m_Connection.Open m_pHelper.GetDSN(strConnection)
    m_bOpened = True
err1:
    If Err.Number <> 0 Then
        m_bOpened = False
        '     MsgBox "Cannot open SQL Server DB: " & Err.Description
    End If
End Sub
Public Sub CloseDB()
    If m_bOpened Then
        m_Connection.Close
        m_bOpened = False
    End If
End Sub



Private Sub m_pCalendar_DoCreateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, _
                                      NewEventID As Long, _
                                      bResult As Boolean)
    bResult = False
    If Not m_bOpened Then: Exit Sub
    'On Error GoTo err1
    Dim strSQL                               As String
    Dim rsNewID                              As ADODB.Recordset
    Dim rsEvent                              As New ADODB.Recordset

    strSQL = "SELECT * FROM CRIS_CalendarEvents WHERE AppointmentID = 0 And AppointmentType=" & ApType
    rsEvent.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
    m_Connection.BeginTrans

Debug.Assert rsEvent.BOF And rsEvent.EOF
    rsEvent.AddNew
    m_pHelper.PutEventToRS pEvent, rsEvent, ApType
    rsEvent.Update
    rsEvent.Close
    m_Connection.CommitTrans
    strSQL = "SELECT MAX(AppointmentID) AS NewEventID FROM CRIS_CalendarEvents"
    Set rsNewID = m_Connection.Execute(strSQL)
    NewEventID = rsNewID("NewEventID")
    rsNewID.Close
    bResult = True
    Exit Sub

    'err1:
    'If Err.Number <> 0 Then
    'Debug.Print "Cannot CreateEvent in DB: " & Err.Description
    'Debug.Assert False
    'End If
End Sub

Private Sub m_pCalendar_DoDeleteEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    If Not m_bOpened Then
        Exit Sub
    End If

    Dim strSQL                               As String

    strSQL = "DELETE FROM CRIS_CalendarEvents WHERE AppointmentID = " & pEvent.ID

    m_Connection.Execute strSQL, , adExecuteNoRecords

    bResult = True
End Sub


Private Sub m_pCalendar_DoRetrieveDayEvents(ByVal dtDay As Date, ByVal pEvents As XtremeCalendarControl.CalendarEvents)
    If Not m_bOpened Then: Exit Sub
    Dim strSQL                               As String
    Dim rsEvents                             As ADODB.Recordset
    Dim pEvent                               As CalendarEvent

    strSQL = m_pHelper.MakeRetrieveDayEventsSQL(dtDay)
    strSQL = strSQL & " And AppointmentType=" & ApType

    Set rsEvents = m_Connection.Execute(strSQL)
    While Not rsEvents.EOF
        Set pEvent = m_pHelper.CreateEventFromRS(rsEvents)
        If Not pEvent Is Nothing Then
            pEvents.Add pEvent
        End If
        rsEvents.MoveNext
    Wend
End Sub

Private Sub m_pCalendar_DoUpdateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    If Not m_bOpened Then: Exit Sub
    Dim rsEvent                              As New ADODB.Recordset
    Dim strSQL                               As String
    bResult = False
    On Error GoTo err1
    strSQL = "SELECT * FROM CRIS_CalendarEvents WHERE AppointmentID = " & pEvent.ID
    rsEvent.Open strSQL, m_Connection, adOpenStatic, adLockOptimistic
    m_Connection.BeginTrans
    m_pHelper.PutEventToRS pEvent, rsEvent, ApType
    rsEvent.Update
    rsEvent.Close
    m_Connection.CommitTrans
    bResult = True
    Exit Sub
err1:
    If Err.Number <> 0 Then
Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
Debug.Assert False
    End If
End Sub
Private Sub Class_Terminate()
    CloseDB
End Sub

