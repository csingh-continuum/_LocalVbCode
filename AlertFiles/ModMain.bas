Attribute VB_Name = "ModMain"

'   Updated by      :Alpa Bandekar
'   Updated On      :08-Feb-2013
'   Purpose         :Changes done Implemented Standard coding format
'                    suspension database changed (AlertConfigDB)
'                    In CloseAlert Added the functionality of AutocloseRefTicketId When Alert is AutoClose. Jira RMM-911
'   VerSion         :1.0.0.2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public cnWPS As New ADODB.Connection
Public CnNOCBO As New ADODB.Connection
Public CnAlertGen As New ADODB.Connection
Public CnAlertSync As ADODB.Connection
Public CnAlertConfig As ADODB.Connection

Public NOCBOdbstr As String
Public ITSWebpostStatusDBstr As String
Public ITSAlertSyncdbstr As String
Public ITSAlertConfigdbstr As String
Public ITSAlertGendbstr As String
Dim constINIFile As String
Public strConTO  As String
Public strCmdTO  As String
Public lngNewTickets As Long
Public lngTicketsClosed As Long

Public Const JOBCATEGORY As String = "ITSAgentMon\MSMA\Alertfiles pending"
Public Const TKT_CATEGORY As String = "PIDMNCHECKSTATUS"

Sub Main()
10        On Error GoTo ERR
         
20        If App.PrevInstance Then
30            End
40        End If
50        CreateDir App.Path & "\Log\"
60        constINIFile = App.Path & "\Ini\dbstring.Ini"
70        CreateLogFile
80        LogNotify "Main", "INFO", "START"
          
90        If OpenConnection = True Then
              LogNotify "Main", "INFO", "Connection Success"
100           AlertFilesNewUpdateClose
110           LogNotify "Main", "INFO", "Total new tickets raised : " & lngNewTickets
120           LogNotify "Main", "INFO", "Total tickets closed : " & lngTicketsClosed
130           CloseConnection
140       End If
150       LogNotify "Main", "INFO", "END"
160       End
170       Exit Sub
ERR:
180       LogNotify "Main", "ERROR", ERR.Number, ERR.Description, Erl
190       LogNotify "Main", "INFO", "END"
200       End
End Sub

Public Sub AlertFilesNewUpdateClose()
    Dim rs As New ADODB.Recordset
    Dim tempJobId As Double
    Dim sJobDescription As String
    Dim sComment As String
    Dim IsError As Boolean
    Dim TransStart As Boolean
    Dim CmdTicketIns As New ADODB.Command
    Dim CmdTicketUpdt As New ADODB.Command
    Dim CmdTicketClose As New ADODB.Command
    Dim RESName As String
    
10       On Error GoTo Er
 
20      Set rs = cnWPS.Execute("USP_PIDMNCheckStatus_Tkt")
30      With rs
40      If .EOF <> True Or .BOF <> True Then
50      While .EOF <> True
60        IsError = False
70        TransStart = True
80        cnWPS.BeginTrans
90        CnNOCBO.BeginTrans
100       CnAlertGen.BeginTrans
110       CnAlertSync.BeginTrans
120       LogNotify "AlertFilesNewUpdateClose", "INFO", "Process Start For RegId : " & !Regid & " AlertId : " & !TktId & ""
130       If DeviceDownExists(!Regid) = False Then
140           tempJobId = IIf(IsNull(!TktId), 0, !TktId)
150           SuspToDate = "12:00:00 AM"
160           Suspended = GetSuspStatus(!Regid, CONDITIONID, "", "", SuspToDate)
170           If Suspended = "N" Then
180                     RESName = ""
190                     If Trim(!ResFriendlyName) <> "" Then
200                           RESName = !ResourceName & "(" & !ResFriendlyName & ")"
210                     Else
220                           RESName = !ResourceName
230                     End If
                           
240               If tempJobId = 0 And !RegIDCount > 0 Then
250                   sJobDescription = ![Desc]
260                   fnCreateNewJob ![Membercode], CONDITIONID, "Alert files pending", sJobDescription, 3, JOBCATEGORY, 2, 0, 0, Format(Now(), "MM\/DD\/YYYY HH:NN:SS AMPM"), ![SiteCode], ![SiteSubCode], !Regid, tempJobId, "", "", RESName, "SAM", "PMT", Format(Now(), "MM\/DD\/YYYY HH:NN:SS AMPM"), !SiteCode & "", "", "", "", 0, 1, Suspended, SuspToDate
270                   If tempJobId <> 0 Then
280                       LogNotify "AlertFilesNewUpdateClose", "INFO", "New Job Creation - SUCCESS. JobId =" & CStr(tempJobId)
                        
290                       Set CmdTicketIns = Nothing
300                       With CmdTicketIns
310                           .ActiveConnection = cnWPS
320                           .CommandType = adCmdStoredProc
330                           .CommandText = "USP_PIDMNCheckEVT_Tkt_IUD"
340                           .CommandTimeout = cnWPS.CommandTimeout
350                           .Parameters.Append .CreateParameter("@INRegID", adBigInt, adParamInput, , rs!Regid)
360                           .Parameters.Append .CreateParameter("@INTktid", adBigInt, adParamInput, , tempJobId)
370                           .Parameters.Append .CreateParameter("@InCategory", adVarChar, adParamInput, 100, TKT_CATEGORY)
380                           .Parameters.Append .CreateParameter("@InOpt", adInteger, adParamInput, , 1)
390                           .Parameters.Append .CreateParameter("@Outret", adInteger, adParamOutput)
400                           .Execute
410                            If .Parameters("@Outret").Value = 1 Then
420                                 LogNotify "AlertFilesNewUpdateClose", "INFO", "New Status Updation Success For AlerID ", tempJobId
430                            Else
440                                 IsError = True
450                                 LogNotify "AlertFilesNewUpdateClose", "ERROR", "New Status Updation Failed For AlerID ", tempJobId
460                            End If
470                       End With
                     
480                       lngNewTickets = lngNewTickets + 1
490                       If GetEscMail(!Regid, CONDITIONID, "", "") <> "" Then
500                            EscMail_IU !Regid, CONDITIONID, "", "", CStr(tempJobId), SuspToDate, "N", "Alert files pending", sJobDescription, "Notify@itsupport247.net", 1, ""
510                       End If
520                   Else
530                       IsError = True
540                       LogNotify "AlertFilesNewUpdateClose", "ERROR", "New Job Creation - FAILED"
550                   End If
560               ElseIf tempJobId > 0 And !RegIDCount > 0 Then
570                   sComment = IIf(IsNull(![Desc]), "", ![Desc])
580                   If UpdateJob(![TktId], sComment, 0, !SiteCode, ![SiteSubCode], !Membercode) Then
590                       LogNotify "AlertFilesNewUpdateClose", "INFO", "Job Update SUCESS"
600                       Set CmdTicketUpdt = Nothing
610                       With CmdTicketUpdt
620                           .ActiveConnection = cnWPS
630                           .CommandType = adCmdStoredProc
640                           .CommandText = "USP_PIDMNCheckEVT_Tkt_IUD"
650                           .CommandTimeout = cnWPS.CommandTimeout
660                           .Parameters.Append .CreateParameter("@INRegID", adBigInt, adParamInput, , rs!Regid)
670                           .Parameters.Append .CreateParameter("@INTktid", adBigInt, adParamInput, , tempJobId)
680                           .Parameters.Append .CreateParameter("@InCategory", adVarChar, adParamInput, 100, TKT_CATEGORY)
690                           .Parameters.Append .CreateParameter("@InOpt", adInteger, adParamInput, 1, 1)
700                           .Parameters.Append .CreateParameter("@Outret", adInteger, adParamOutput)
710                           .Execute
720                           If .Parameters("@Outret").Value = 1 Then
730                                LogNotify "AlertFilesNewUpdateClose", "INFO", "USP_PIDMNCheckEVT_Tkt_IUD Success for Update AlertID ", tempJobId
740                           Else
750                                IsError = True
760                                LogNotify "AlertFilesNewUpdateClose", "ERROR", "USP_PIDMNCheckEVT_Tkt_IUD Failed for Update AlertID ", tempJobId
770                           End If
780                       End With
790                   Else
800                       IsError = True
810                       LogNotify "AlertFilesNewUpdateClose", "ERROR", "Job Update FAILED for AlertId: ", tempJobId
820                   End If
830               ElseIf tempJobId > 0 And !RegIDCount = 0 Then
840                   If UpdateJob(![TktId], "System Closed this Job, based on alert sent", 1, !SiteCode, ![SiteSubCode], !Membercode) Then
850                       LogNotify "AlertFilesNewUpdateClose", "INFO", "Job Close SUCCESS"
860                        Set CmdTicketClose = Nothing
870                        With CmdTicketClose
880                           .ActiveConnection = cnWPS
890                           .CommandType = adCmdStoredProc
900                           .CommandText = "USP_PIDMNCheckEVT_Tkt_IUD"
910                           .CommandTimeout = cnWPS.CommandTimeout
920                           .Parameters.Append .CreateParameter("@INRegID", adBigInt, adParamInput, , rs!Regid)
930                           .Parameters.Append .CreateParameter("@INTktid", adBigInt, adParamInput, , tempJobId)
940                           .Parameters.Append .CreateParameter("@InCategory", adVarChar, adParamInput, 100, TKT_CATEGORY)
950                           .Parameters.Append .CreateParameter("@InOpt", adInteger, adParamInput, , 2)
960                           .Parameters.Append .CreateParameter("@Outret", adInteger, adParamOutput)
970                           .Execute
980                            If .Parameters("@Outret").Value = 1 Then
990                                 LogNotify "AlertFilesNewUpdateClose", "INFO", "USP_PIDMNCheckEVT_Tkt_IUD Success for Close AlertID ", tempJobId
                                                           'Calling Noc sp to Close reference tickets
1000                                If fnCloseRefTkt(CStr(tempJobId), "AGENTMON", App.EXEName, 5) = False Then
1010                                    LogNotify "AlertFilesNewUpdateClose", "ERROR", " USP_Close_RefTicketId Failed for Close RefTkt AlertID " & CStr(tempJobId)
1020                                Else
1030                                    LogNotify "AlertFilesNewUpdateClose", "INFO", " USP_Close_RefTicketId Success for Close RefTkt AlertID " & CStr(tempJobId)
1040                                End If

1050                           Else
1060                                IsError = True
1070                                LogNotify "AlertFilesNewUpdateClose", "ERROR", "USP_PIDMNCheckEVT_Tkt_IUD Failed for Close AlertID ", tempJobId
1080                           End If
1090                      End With

1100                      lngTicketsClosed = lngTicketsClosed + 1
1110                      If GetEscMail(CLng(!Regid), CLng(CONDITIONID), "", "") <> "" Then
1120                          EscMail_IU CLng(!Regid), CLng(CONDITIONID), "", "", CStr(tempJobId), SuspToDate, "C", "Alert files pending - Issue resolved", "System Closed this Job, based on alert sent", "Notify@itsupport247.net", 1, ""
1130                      End If
1140                  Else
1150                      IsError = True
1160                      LogNotify "AlertFilesNewUpdateClose", "ERROR", "Job Close FAILED for AlertId: ", tempJobId
1170                  End If
1180              End If 'Tempjobid condition
1190          End If   'Suspended Y/N
1200      Else
1210          LogNotify "AlertFilesNewUpdateClose", "INFO", "Devicedown Alert found"
1220      End If
          
1230      If TransStart = True And IsError = False Then
1240            TransStart = False
1250            cnWPS.CommitTrans
1260            CnNOCBO.CommitTrans
1270            CnAlertGen.CommitTrans
1280            CnAlertSync.CommitTrans
1290            LogNotify "AlertFilesNewUpdateClose", "INFO", "Transaction Committed"
1300      ElseIf TransStart = True And IsError = True Then
1310            TransStart = False
1320            cnWPS.RollbackTrans
1330            CnNOCBO.RollbackTrans
1340            CnAlertGen.RollbackTrans
1350            CnAlertSync.RollbackTrans
1360            LogNotify "AlertFilesNewUpdateClose", "ERROR", "Transaction Rollbacked"
1370      End If
NextRec:
1380  .MoveNext
1390  Wend
1400    Else
1410        LogNotify "AlertFilesNewUpdateClose", "INFO", "No Record Found to Process"
1420    End If
1430      End With
    
1440       If Not CmdTicketIns Is Nothing Then Set CmdTicketIns = Nothing
1450       If Not CmdTicketUpdt Is Nothing Then Set CmdTicketUpdt = Nothing
1460       If Not CmdTicketClose Is Nothing Then Set CmdTicketClose = Nothing
1470  Exit Sub
Er:

1480      LogNotify "AlertFilesNewUpdateClose", "ERROR", ERR.Number, ERR.Description, Erl
1490      If TransStart = True Then
1500          TransStart = False
1510            cnWPS.RollbackTrans
1520            CnNOCBO.RollbackTrans
1530            CnAlertGen.RollbackTrans
1540            CnAlertSync.RollbackTrans
1550          LogNotify "AlertFilesNewUpdateClose", "ERROR", "Transaction Rollbacked"
1560      End If
1570      If Not rs Is Nothing Then
1580          If rs.State = 1 Then
1590              If Not rs.EOF Then
1600                  LogNotify "AlertFilesNewUpdateClose", "ERROR", "We are skipping this RegId ", rs("RegId") & ""
1610                  Resume NextRec
1620              End If
1630          End If
1640          Set rs = Nothing
1650      End If
1660      If Not CmdTicketIns Is Nothing Then Set CmdTicketIns = Nothing
1670      If Not CmdTicketUpdt Is Nothing Then Set CmdTicketUpdt = Nothing
1680      If Not CmdTicketClose Is Nothing Then Set CmdTicketClose = Nothing
End Sub
 

Function CloseConnection()
10        On Error Resume Next
20        If Not CnAlertGen Is Nothing Then
30            If CnAlertGen.State = 1 Then CnAlertGen.Close
40            Set CnAlertGen = Nothing
50        End If

60        If Not CnAlertSync Is Nothing Then
70            If CnAlertSync.State = 1 Then CnAlertSync.Close
80            Set CnAlertSync = Nothing
90        End If

100       If Not CnAlertConfig Is Nothing Then
110           If CnAlertConfig.State = 1 Then CnAlertConfig.Close
120           Set CnAlertConfig = Nothing
130       End If

140       If Not cnWPS Is Nothing Then
150           If cnWPS.State = 1 Then cnWPS.Close
160           Set cnWPS = Nothing
170       End If

180       If Not CnNOCBO Is Nothing Then
190           If CnNOCBO.State = 1 Then CnNOCBO.Close
200           Set CnNOCBO = Nothing
210       End If
 End Function

Function OpenConnection() As Boolean
10        On Error GoTo errHandle
20        If getConDtls Then
          
30            If Not getConnected(CnNOCBO, NOCBOdbstr) Then
40                LogNotify "OpenConnection", "ERROR", "Could not connect to NOCBO Database"
50                OpenConnection = False
60                Exit Function
70            End If
80            If Not getConnected(cnWPS, ITSWebpostStatusDBstr) Then
90                LogNotify "OpenConnection", "ERROR", "Could not connect to ITSWebpostStatusdb Database"
100               OpenConnection = False
110               Exit Function
120           End If
              
130           If Not getConnected(CnAlertGen, ITSAlertGendbstr) Then
140              LogNotify "OpenConnection", "ERROR", "Could not connect to ITSAlertGenerator Database"
150              OpenConnection = False
160              Exit Function
170           End If
        
180           If Not getConnected(CnAlertSync, ITSAlertSyncdbstr) Then
190              LogNotify "OpenConnection", "ERROR", "Could not connect to ITSAlertSyncdb  Database"
200              OpenConnection = False
210              Exit Function
220           End If

230           If Not getConnected(CnAlertConfig, ITSAlertConfigdbstr) Then
240              LogNotify "OpenConnection", "ERROR", "Could not connect to ITSAlertConfigdb Database"
250              OpenConnection = False
260              Exit Function
270           End If



280           OpenConnection = True
290       Else
300           OpenConnection = False
310       End If
320       Exit Function
errHandle:
330       OpenConnection = False
340       LogNotify "OpenConnection", "ERROR", ERR.Number, ERR.Description, Erl
End Function

Function getConDtls() As Boolean
10        On Error GoTo errHandle
          
20        NOCBOdbstr = getProfVals("AlertFiles", "NOCBOdbstr", constINIFile)
30        ITSWebpostStatusDBstr = getProfVals("AlertFiles", "ITSWebpostStatusDBstr", constINIFile)
40        ITSAlertSyncdbstr = getProfVals("AlertFiles", "ITSAlertSyncdbstr", constINIFile)
50        ITSAlertGendbstr = getProfVals("AlertFiles", "ITSAlertGendbstr", constINIFile)
60        ITSAlertConfigdbstr = getProfVals("AlertFiles", "ITSAlertConfigdbstr", constINIFile)

70        strConTO = getProfVals("AlertFiles", "conntimeoutsec", constINIFile)
80        strCmdTO = getProfVals("AlertFiles", "cmdtimeoutsec", constINIFile)

90        If Trim$(ITSWebpostStatusDBstr) = "" Or Trim$(ITSAlertSyncdbstr) = "" Or Trim$(ITSAlertGendbstr) = "" Or Trim$(ITSAlertConfigdbstr) = "" Or Trim$(NOCBOdbstr) = "" Then
100           LogNotify "getConDtls", "ERROR", "-1", "DBConnection Data missing In AlertFiles"
110           getConDtls = False
120       Else
130           getConDtls = True
140       End If
150       Exit Function
errHandle:
160       LogNotify "getConDtls", "ERROR", ERR.Number, ERR.Description, Erl
170       getConDtls = False
End Function

Function getConnected(ObjConn As Object, conStr As String) As Boolean
10        On Error GoTo ErrWhileConnecting
          Dim strConnStr
20        Set ObjConn = New ADODB.Connection
30        ObjConn.ConnectionString = conStr
40        ObjConn.ConnectionTimeout = strConTO
50        ObjConn.CommandTimeout = strCmdTO
60        ObjConn.CursorLocation = adUseClient
70        ObjConn.Open
80        getConnected = True
90        Exit Function
ErrWhileConnecting:
100       LogNotify "getConnected", "ERROR", ERR.Number, ERR.Description, Erl
110       getConnected = False
120       Exit Function
End Function

