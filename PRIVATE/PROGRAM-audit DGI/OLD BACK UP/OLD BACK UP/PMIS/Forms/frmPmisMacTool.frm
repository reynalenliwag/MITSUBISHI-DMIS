VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPmisMacTool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " MAC TOOL"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmPmisMacTool.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   3675
      Begin VB.CommandButton cmdCompute 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3030
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPmisMacTool.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   30
         TabIndex        =   4
         Top             =   600
         Width           =   3585
         Begin MSComctlLib.ProgressBar progCPB 
            Height          =   285
            Left            =   30
            TabIndex        =   5
            Top             =   120
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Max             =   1
            Scrolling       =   1
         End
      End
      Begin VB.Label labCPB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
      Begin VB.Label labPartnumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   270
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmPmisMacTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompute_Click()


    Dim rsAlldaytran As New ADODB.Recordset
    Dim rsDAYTRAN As New ADODB.Recordset
    Dim rsTdayTran As New ADODB.Recordset
    Dim rsStockno As New ADODB.Recordset
    Dim rsYear As New ADODB.Recordset
    Dim rsMonth As New ADODB.Recordset
    Dim rsPartno As New ADODB.Recordset
    Dim rsAllpartno As New ADODB.Recordset
    
    Dim blnNoMonthChange As Boolean
    
    Dim TranType As String
    Dim NEWMAC As Currency
    Dim TOTALQTY As Integer
    Dim maxYear As Integer
    Dim minYear As Integer
    Dim maxMonth As Integer
    Dim minMonth As Integer
    Dim intMonth As Integer
    Dim intYear As Integer
    Dim prevMonth As Integer
    Dim NEWMONTH As Integer
    Dim RivCtr As Integer
    Dim PartCtr As Integer
    Dim strPartnumber As String
    Dim RRCtr As Integer
    Dim AdjCtrIn As Integer
    Dim adjCtrOut As Integer
    Dim I As Integer
    
    'strPartnumber = Trim(txtPartnumber) for manual input
    
    'Set rsAllpartno = gconDMIS.Execute("select * from pmis_stockmas where active = 'Y' and type = 'P'")
    Call rsAllpartno.Open("select * from pmis_stockmas where active = 'Y' and type = 'P'", gconDMIS, adOpenStatic)
    progCPB.Value = 1
    progCPB.Max = rsAllpartno.RecordCount
        
    If Not (rsAllpartno.BOF And rsAllpartno.EOF) Then
        Do While Not rsAllpartno.EOF
            labPartnumber.Caption = rsAllpartno!STOCKNO
                DoEvents
            
            strPartnumber = Trim(rsAllpartno!STOCKNO)
            'If strPartnumber = "08170-2E000" Then Stop
            
           
            blnNoMonthChange = False
            'maxMonth = 12
        
            gconDMIS.Execute ("update pmis_stockmas set lastm_mac = 0, lastm_oh = 0, mac = 0, onhand = 0 where stockno = '" & strPartnumber & "' ")
            NEWMAC = 0
            ' get max and minimun year of trandate of the partnumber
            Set rsYear = gconDMIS.Execute("select max(year(trandate)) as mxYear, min(year(trandate)) as mnYear from pmis_alldaytran where stock_ord = '" & strPartnumber & "' ")
            
            If Not (IsNull(rsYear!mxYear) And IsNull(rsYear!mnYear)) Then
            maxYear = rsYear!mxYear: minYear = rsYear!mnYear
                'compute whole year transaction of partnumber
                For intYear = minYear To maxYear
                    'If intYear = "2008" Then Stop
                
                    'compute whole month per year
                    Set rsMonth = gconDMIS.Execute("select Max(Month(trandate)) AS mxMonth, Min(Month(trandate)) as mnMonth  from pmis_alldaytran where stock_ord = '" & strPartnumber & "' and year(trandate) = " & intYear & " ")
                   
                    If Not (IsNull(rsMonth!mxMonth) And IsNull(rsMonth!mnMonth)) Then
                        For minMonth = 1 To rsMonth!mxMonth
                            'If minMonth = 3 Then Stop
                             RRCtr = 0: AdjCtrIn = 0: adjCtrOut = 0
                            
                            'update lastm_oh and lastm_mac every end of the month
                            If blnNoMonthChange Then
                                    If prevMonth <> minMonth Then
                                        Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                        TOTALQTY = (rsPartno!ONHAND)
                                        gconDMIS.Execute (" update pmis_stockmas set lastm_mac = " & rsPartno!Mac & ", lastm_oh = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                    End If
                                    blnNoMonthChange = False
                            End If
                        
                        
                             blnNoMonthChange = False
                             Set rsAlldaytran = gconDMIS.Execute("select id, type, trantype, tranno, stock_ord, tranqty,  tranucost, status, in_out, Mac, trandate from pmis_alldaytran where trantype in('RR','RIV','CSH','CHG','ADB','DR','BEG','ADJ') and stock_ord = '" & strPartnumber & "' and status not IN('N','C') and year(trandate) = " & intYear & " and month(trandate) = " & minMonth & " order by trandate asc, id asc,tranno asc ")
                             
                             If Not (rsAlldaytran.BOF And rsAlldaytran.EOF) Then
                             Do While Not rsAlldaytran.EOF
                             'prevMonth = minMonth
                             blnNoMonthChange = True
                             'If rsAlldaytran!TranType = "RR" Then Stop
                             
                                                        
                                        '==[get the beggining mac]==
                                        If Left((rsAlldaytran!TranType), 3) = "BEG" Then
                                            If IsNull(rsAlldaytran!Mac) Then
                                                NEWMAC = 0
                                                NEWMAC = Round(rsAlldaytran!TRANUCOST, 2)
                                                    Set rsDAYTRAN = gconDMIS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                            gconDMIS.Execute ("update pmis_daytran set  mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & "  ")
                                                    End If
                                                       
                                                    Set rsTdayTran = gconDMIS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & "   ")
                                                    If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                            gconDMIS.Execute ("update pmis_tdaytran set mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    End If
                                            Else
                                                NEWMAC = 0
                                                NEWMAC = Round(rsAlldaytran!Mac, 2)
                                                gconDMIS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand  = " & rsAlldaytran!tranqty & "  where stockno = '" & strPartnumber & "'  ")
                                            End If
                                       
                                        ' ==[compute for the new mac]==
                                       ElseIf (rsAlldaytran!TranType) = "RR" Then
                                            RRCtr = RRCtr + 1
                                            Set rsStockno = gconACCESS.Execute("select stockno, lastm_mac, lastm_oh, mac, onhand from pmis_stockmas where stockno in (select stock_ord from pmis_alldaytran where stock_ord = '" & strPartnumber & "')")
                                            '
                                            If (rsStockno!lastm_oh) > 0 And rsStockno!ONHAND > 0 Then
                                            
                                                If (rsStockno!lastm_oh) = rsStockno!ONHAND Then
                                                    
                                                    If RRCtr = 1 Then
                                                    
                                                        TOTALQTY = (rsStockno!lastm_oh + rsAlldaytran!tranqty)
                                                        
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!lastm_mac * rsStockno!lastm_oh) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & " where stockno =  '" & strPartnumber & "'  ")
                                                    Else
                                                    TOTALQTY = (rsStockno!lastm_oh + rsAlldaytran!tranqty)
                                                        
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!Mac * rsStockno!lastm_oh) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & " where stockno =  '" & strPartnumber & "'  ")
                                                    
                                                    End If
                                                    
                                                        Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                                gconACCESS.Execute ("update pmis_daytran set  mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                        Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                                gconACCESS.Execute ("update pmis_tdaytran set mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        End If
                                                Else
                                                    If RRCtr <> 1 Then
                                                        TOTALQTY = (rsStockno!ONHAND + rsAlldaytran!tranqty)
                                                        
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!Mac * rsStockno!ONHAND) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & "  where stockno = '" & strPartnumber & "' ")
                                                    Else
                                                        TOTALQTY = (rsStockno!ONHAND + rsAlldaytran!tranqty)
                                                        
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!lastm_mac * rsStockno!ONHAND) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & "  where stockno = '" & strPartnumber & "' ")
                                                    End If
                                                    
                                                    If (rsStockno!lastm_oh) <> rsStockno!ONHAND Then
                                                        Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                                gconACCESS.Execute ("update pmis_daytran set  mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                        Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                                gconACCESS.Execute ("update pmis_tdaytran set mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                    End If
                                                End If
                                            'if theres an adjustment in quantity
                                            ElseIf (rsStockno!ONHAND) > 0 And (rsStockno!lastm_oh) = 0 Then
                                                    If RRCtr = 1 Then
                                                        TOTALQTY = (rsStockno!ONHAND + rsAlldaytran!tranqty)
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!lastm_mac * rsStockno!ONHAND) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & " where stockno = '" & strPartnumber & "'  ")
                                                    Else
                                                        TOTALQTY = (rsStockno!ONHAND + rsAlldaytran!tranqty)
                                                        NEWMAC = 0
                                                        NEWMAC = Round((((rsStockno!Mac * rsStockno!ONHAND) + (rsAlldaytran!TRANUCOST * rsAlldaytran!tranqty)) / TOTALQTY), 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand = " & RivCtr & " where stockno = '" & strPartnumber & "'  ")
                                                    End If
                                                    
                                                    Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                            gconACCESS.Execute ("update pmis_daytran set  mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        End If
                                                    Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                            gconACCESS.Execute ("update pmis_tdaytran set mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                        End If
                                            Else
                                            'onhand is zero
                                                    If (rsStockno!ONHAND) = 0 Then
                                                        NEWMAC = 0
                                                        NEWMAC = Round(rsAlldaytran!TRANUCOST, 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand  = " & RivCtr & " where stockno = '" & strPartnumber & "'  ")
                                        
                                                            Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                                gconACCESS.Execute ("update pmis_daytran set  mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                        Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                                gconACCESS.Execute ("update pmis_tdaytran set mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                            ' onhand is negative
                                                    ElseIf (rsStockno!ONHAND) < 0 Then
                                                        NEWMAC = 0
                                                        NEWMAC = Round(rsAlldaytran!TRANUCOST, 2)
                                                        RivCtr = rsAlldaytran!tranqty + rsStockno!ONHAND
                                                        gconACCESS.Execute (" update pmis_stockmas set mac = " & NEWMAC & ", onhand  = " & RivCtr & " where stockno = '" & strPartnumber & "'   ")
                                                    End If
                                             End If
                                       End If
                                         
                                        '==[ update the mac and tranucost in daytran or tdayran if there's a issuance ]===
                                        If rsAlldaytran!TranType = "RIV" Or rsAlldaytran!TranType = "CSH" Or rsAlldaytran!TranType = "CHG" Or rsAlldaytran!TranType = "DR" Then
                                            Set rsDAYTRAN = gconDMIS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                    gconDMIS.Execute ("update pmis_daytran set  tranucost = '" & NEWMAC & "', mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    RivCtr = rsAlldaytran!tranqty
                                                End If
                                                Set rsTdayTran = gconDMIS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                    gconDMIS.Execute ("update pmis_tdaytran set  tranucost = '" & NEWMAC & "', mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    RivCtr = rsAlldaytran!tranqty
                                                End If
                                        ' update the tranucost in Advance bill transaction
                                        ElseIf rsAlldaytran!TranType = "ADB" Then
                                                Set rsDAYTRAN = gconDMIS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                    gconDMIS.Execute ("update pmis_daytran set  tranucost = '" & NEWMAC & "', mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    RivCtr = 0
                                                End If
                                                Set rsTdayTran = gconDMIS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                    gconDMIS.Execute ("update pmis_tdaytran set  tranucost = '" & NEWMAC & "', mac = '" & NEWMAC & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                    RivCtr = rsAlldaytran!tranqty
                                                End If
                                        End If
                                                            'Next
                                         '==[ update onhand when there is  a issuance transaction ]==
                                        If rsAlldaytran!TranType = "RIV" Or rsAlldaytran!TranType = "CSH" Or rsAlldaytran!TranType = "CHG" Or rsAlldaytran!TranType = "DR" Then
                                            Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                               TOTALQTY = (rsPartno!ONHAND - RivCtr)
                                            gconDMIS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                        '===[ ADJUSTMENT TRANSACTION ] ===
                                        ElseIf rsAlldaytran!TranType = "ADJ" And rsAlldaytran!IN_OUT = "I" Then
                                                AdjCtrIn = 1 + adjCtrOut
                                                
                                                If AdjCtrIn = 1 Then
                                                Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                                    TOTALQTY = (rsPartno!lastm_oh + rsAlldaytran!tranqty)
                                        
                                                        If rsPartno!ONHAND = 0 And rsPartno!lastm_oh = 0 And rsPartno!lastm_mac = 0 And rsPartno!Mac = 0 Then
                                                            Set rsTdayTran = gconACCESS.Execute("select * from pmis_alldaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & ", mac = " & rsTdayTran!TRANUCOST & " where stockno = '" & strPartnumber & "'")
                                                            NEWMAC = Round(rsTdayTran!TRANUCOST, 2)
                                                        Else
                                                            gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                                        End If
                                                 Else
                                                        Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                                            TOTALQTY = (rsPartno!ONHAND + rsAlldaytran!tranqty)
                                                        If rsPartno!ONHAND = 0 And rsPartno!lastm_oh = 0 And rsPartno!lastm_mac = 0 And rsPartno!Mac = 0 Then
                                                            Set rsTdayTran = gconACCESS.Execute("select * from pmis_alldaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & ", mac = " & rsTdayTran!TRANUCOST & " where stockno = '" & strPartnumber & "'")
                                                            NEWMAC = Round(rsTdayTran!TRANUCOST, 2)
                                                        Else
                                                            gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                                        End If
                                                                 
                                                 
                                                 End If
                                         
                                         ElseIf rsAlldaytran!TranType = "ADJ" And rsAlldaytran!IN_OUT = "O" Then
                                                 adjCtrOut = adjCtrOut + 1
                                                 
                                                 If RRCtr = 1 Then
                                                        Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                                        TOTALQTY = (rsPartno!lastm_oh - rsAlldaytran!tranqty)
                                                        gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                    
                                                        Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                                gconACCESS.Execute ("update pmis_daytran set  mac = '" & rsDAYTRAN!TRANUCOST & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                        Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                                gconACCESS.Execute ("update pmis_tdaytran set mac = '" & rsTdayTran!TRANUCOST & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                  Else
                                                        Set rsPartno = gconACCESS.Execute("select lastm_oh,lastm_mac,mac,onhand from pmis_stockmas where stockno = '" & strPartnumber & "' ")
                                                        TOTALQTY = (rsPartno!ONHAND - rsAlldaytran!tranqty)
                                                        gconACCESS.Execute (" update pmis_stockmas set onhand = " & TOTALQTY & " where stockno = '" & strPartnumber & "'")
                                    
                                                        Set rsDAYTRAN = gconACCESS.Execute("select * from pmis_daytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsDAYTRAN.EOF And rsDAYTRAN.BOF) Then
                                                                gconACCESS.Execute ("update pmis_daytran set  mac = '" & rsDAYTRAN!TRANUCOST & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                        Set rsTdayTran = gconACCESS.Execute("select * from pmis_tdaytran where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            If Not (rsTdayTran.EOF And rsTdayTran.BOF) Then
                                                                gconACCESS.Execute ("update pmis_tdaytran set mac = '" & rsTdayTran!TRANUCOST & "'  where stock_ord = '" & strPartnumber & "' and trantype = '" & rsAlldaytran!TranType & "'  and tranno = " & rsAlldaytran!TRANNO & " ")
                                                            End If
                                                  End If
                                         End If
                                        
                    
                            rsAlldaytran.MoveNext
                        Loop
                    End If
                            
                    prevMonth = minMonth
                Next ' month loop
            End If
        Next ' year loop
    End If
        
            I = I + 1
            progCPB.Value = I
            labCPB.Caption = Round((progCPB.Value / progCPB.Max * 100), 0) & "%"
            DoEvents
    
    
            rsAllpartno.MoveNext
        Loop
    End If 'end of rsAllpartno

    MsgBox "FIXING DONE"

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0

    CenterMe frmMain, Me, 1
    
End Sub

