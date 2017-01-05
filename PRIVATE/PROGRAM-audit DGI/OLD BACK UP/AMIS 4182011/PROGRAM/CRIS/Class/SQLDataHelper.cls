VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_pCalendar                              As CalendarControl
Public Sub SetCalendar(ByVal pCalendar As CalendarControl)
    Set m_pCalendar = pCalendar
End Sub

Private Function RoundNum2D(nNumber As Long) As String
Debug.Assert nNumber < 100
    RoundNum2D = IIf(nNumber < 10, "0", "") & nNumber
End Function


Public Function MakeSQLDateTime(dtDate As Date) As String
    Dim strDate                              As String

    ' "{ts'%Y-%m-%d %H:%M:%S'}"
    strDate = "{ts'" & Year(dtDate) & "-"
    strDate = strDate + RoundNum2D(Month(dtDate)) & "-"
    strDate = strDate + RoundNum2D(Day(dtDate)) & " "

    strDate = strDate + RoundNum2D(Hour(dtDate)) & ":"
    strDate = strDate + RoundNum2D(Minute(dtDate)) & ":"
    strDate = strDate + RoundNum2D(Second(dtDate)) & "'}"

    MakeSQLDateTime = strDate
End Function

Public Function MakeSQLstr(ByVal strVB As String) As String
    strVB = Replace(strVB, "\", "\\")
    strVB = Replace(strVB, """", "\""")

    strVB = """" & strVB & """"

    MakeSQLstr = strVB
End Function

Public Function CreateEventFromRS(pEventRS As ADODB.Recordset) As CalendarEvent
    '    On Error GoTo err1
    Set CreateEventFromRS = Nothing
    Dim pEvent                               As CalendarEvent
    Dim nEventID                             As Long

    If pEventRS.EOF Then
        Exit Function
    End If
    nEventID = pEventRS("AppointmentID")
    Set pEvent = m_pCalendar.DataProvider.CreateEventEx(nEventID)
    If pEvent Is Nothing Then
        Exit Function
    End If
    pEvent.Subject = pEventRS("Data1")
    pEvent.Location = pEventRS("Location")
    pEvent.Body = pEventRS("Body")
    pEvent.MeetingFlag = IIf(pEventRS("IsMeeting") <> 0, True, False)
    pEvent.PrivateFlag = IIf(pEventRS("IsPrivate") <> 0, True, False)
    pEvent.Label = pEventRS("LabelID")
    pEvent.BusyStatus = pEventRS("BusyStatus")
    pEvent.Importance = pEventRS("ImportanceLevel")
    pEvent.StartTime = pEventRS("StartDateTime")
    pEvent.EndTime = pEventRS("EndDateTime")
    pEvent.AllDayEvent = IIf(pEventRS("IsAllDayEvent") <> 0, True, False)
    pEvent.Reminder = IIf(pEventRS("IsReminder") <> 0, True, False)
    'pEvent.ReminderMinutesBeforeStart = pEventRS("ReminderMinutesBeforeStart")
    Set CreateEventFromRS = pEvent
    Exit Function
    'err1:
    'If Err.Number <> 0 Then
    'Debug.Print "Cannot read event from recordset: " & Err.Description
    'Debug.Assert False
    '   End If
End Function

Public Sub PutEventToRS(pEvent As CalendarEvent, pEventRS As ADODB.Recordset, intAptype As Integer)
    'On Error GoTo err1
    
    pEventRS("Data1") = pEvent.Subject
    pEventRS("Location") = pEvent.Location
    pEventRS("Body") = pEvent.Body
    pEventRS("IsMeeting") = IIf(pEvent.MeetingFlag, 1, 0)
    pEventRS("IsPrivate") = IIf(pEvent.PrivateFlag, 1, 0)
    pEventRS("LabelID") = pEvent.Label
    pEventRS("BusyStatus") = pEvent.BusyStatus
    pEventRS("ImportanceLevel") = pEvent.Importance
    pEventRS("StartDateTime") = pEvent.StartTime
    pEventRS("EndDateTime") = pEvent.EndTime
    pEventRS("AppointmentType") = intAptype
'    pEventRS("IsAllDayEvent") = IIf(pEvent.AllDayEvent, 1, 0)
'    pEventRS("IsReminder") = IIf(pEvent.Reminder, 1, 0)
'    pEventRS("ScheduleID") = pEvent.ScheduleID
'
    'err1:
    'If Err.Number <> 0 Then
    'Debug.Print "Cannot put event to recordset: " & Err.Description
    'Debug.Assert False
    '    End If
End Sub
Public Function MakeRetrieveDayEventsSQL(ByVal dtDay As Date) As String
    Dim nYear As Long, nMonth As Long, nDay  As Long
    Dim strSQL                               As String

    nYear = Year(dtDay)
    nMonth = Month(dtDay)
    nDay = Day(dtDay)

    strSQL = "SELECT * FROM CRIS_CalendarEvents WHERE  " & vbCrLf

    strSQL = strSQL + "( YEAR(StartDateTime) < " & nYear & vbCrLf
    strSQL = strSQL + "  OR ( YEAR(StartDateTime) = " & nYear & " AND " & vbCrLf
    strSQL = strSQL + "       ( MONTH(StartDateTime) < " & nMonth & " OR "
    strSQL = strSQL + "         MONTH(StartDateTime) = " & nMonth & " AND " & vbCrLf
    strSQL = strSQL + "         DAY(StartDateTime) <= " & nDay & vbCrLf

    strSQL = strSQL + "     ) ) )AND " & vbCrLf

    strSQL = strSQL + "( YEAR(EndDateTime) > " & nYear & vbCrLf
    strSQL = strSQL + "  OR ( YEAR(EndDateTime) = " & nYear & " AND " & vbCrLf
    strSQL = strSQL + "       (MONTH(EndDateTime) > " & nMonth & " OR "
    strSQL = strSQL + "        MONTH(EndDateTime) = " & nMonth & " AND " & vbCrLf
    strSQL = strSQL + "        DAY(EndDateTime) >= " & nDay & vbCrLf
    strSQL = strSQL + "     ) ) ) " & vbCrLf
    MakeRetrieveDayEventsSQL = strSQL
End Function
Public Function GetDSN(strConnection As String)
    GetDSN = Right(strConnection, Len(strConnection) - Len("Provider=Custom;"))
End Function


