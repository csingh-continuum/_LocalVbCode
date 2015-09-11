Attribute VB_Name = "ModFunction"
Option Explicit

Public Function fnCreateNewJob(strGroupName As String, lngCondID As Long, strJobName As String, strJobDescription As String, Optional intJobType As Integer, Optional strJobCategoryName As String, Optional intCategoryType As Integer, Optional intJobStatus As Integer, Optional intPublish As Integer, Optional strStartDate As Date, Optional strClient As String, Optional strLocation As String, Optional intRegID As Long, Optional lngJobID As Double, Optional strCompletionDate As String, Optional strCompletionDuration As String, Optional strResource As String, Optional strModule As String, Optional strAlertGroup As String, Optional strMNDateTime As Date, Optional strThresholdCounter As String, Optional strThresholdValue As String, Optional strThresholdStatus As String, Optional strAdditionalDetails As String, Optional intTaskID As Integer, Optional intVer As Integer, Optional INSuspensionFlag As String, Optional INDATETime As Date)
      'Function fnCreateNewJob : This function will create job in NOC@SAAZ.
      '       : This is for USIDC only.
      '------------------------------------------------------------------------------------------------------------------------------
      ' Description   :
      '------------------------------------------------------------------------------------------------------------------------------
      '  Byval strGroupName,   : Management Node Group Name
      '  Byval strJobName,   : Job Name
      '  Byval strJobDescription, : Job Description
      '  Byval intJobType,   : 1 - Auto-generated Non closed, 2 -Manually generated Non Closed
      '         : 3 - Auto-generated Closed, 4 -Manually generated Closed
      '  Byval strJobCategoryName, : Job category Name like Server\Backup\Veritas
      '  Byval intCategoryType,  : 0 : Manual, 1 - System, 2- : System for System Close.
      '  Byval intJobStatus,   : 0 : New, 3 : Forced Close, 4 : Close, 5 : Auto Close, 6 : Discard,
      '  Byval intPublish,   : 0 : Not publish to client , 1 :  Publish to Client
      '  byval strStartDate,   : Job added or job start date
      '  Byval strCompletionDate, : Job Completion Date
      '  Byval strCompletionDuration, : Job Completion duration in hours if any
      '  Byval strClient,   : MN Client Name
      '  Byval strLocation,   : MN Location
      '  Byval strResource,   : Mn resource from this job is generated
      '  Byval strModule,   : MN Module name from which this job is generated
      '  Byref lngJobID    : > JobID after sucessfully generated, 0: if error occured.
      '  byval intVer                  :by defalut version 1
      '------------------------------------------------------------------------------------------------------------------------------
10    On Error GoTo bye
       
          Dim CMDJobMgmt As ADODB.Command
          Dim lngMaxJobID As Long
          Dim lngAssignTo As Long
20        Set CMDJobMgmt = New ADODB.Command
         '=1 intJobType
30        lngMaxJobID = 0
40        intPublish = 0
50        strAlertGroup = ""
60        strThresholdCounter = ""
70        strThresholdValue = ""
80        strThresholdStatus = ""
90        strAdditionalDetails = ""
100       intTaskID = 0
110       strCompletionDuration = ""
120       If Len(strCompletionDate) <= 0 Then
130          strCompletionDate = ""
140       End If
150       If Len(strCompletionDuration) <= 0 Then
160          strCompletionDuration = ""
170       End If
180       If Len(intPublish) <= 0 Then
190          intPublish = 0
200       End If
210       If Len(lngAssignTo) <= 0 Then
220          lngAssignTo = 0
230       End If
          
240      If Len(intTaskID) <= 0 Then
250         intTaskID = 0
260      End If
         
270      With CMDJobMgmt
280         .ActiveConnection = CnNOCBO
290         .CommandText = "JobManagement_sod"
300         .CommandType = 4
310         .Parameters.Append .CreateParameter("Return", 5, 4) 'Return parameter
            
320         .Parameters.Append .CreateParameter("MSPName", 200, 1, 100, strGroupName) 'SDMTEST
330         .Parameters.Append .CreateParameter("JobName", 200, 1, 255, strJobName) ' Antivirus Not Supported Systems
340         .Parameters.Append .CreateParameter("JobDescription", 203, 1, Len(strJobDescription) + 1, strJobDescription) ' <JobDescription><SystemCount><![CDATA[ Antivirus Not Supported [System Count = 1] ]]></SystemCount></JobDescription>
350         .Parameters.Append .CreateParameter("JobTypeID", 3, 1, , intJobType)     ' 1
360         .Parameters.Append .CreateParameter("CategoryName", 200, 1, 5000, strJobCategoryName)  ' Prev.Maint\AVNotSupported\Desktops
370         .Parameters.Append .CreateParameter("intCatType", 3, 1, , intCategoryType)  ' 1
380         .Parameters.Append .CreateParameter("intStatusID", 3, 1, , intJobStatus)  '0
390         .Parameters.Append .CreateParameter("Ispublish", 3, 1, , Null) 'intPublish)  ' 0
400         .Parameters.Append .CreateParameter("StartDateTime", 134, 1, , strStartDate) ' 06/09/18
410         .Parameters.Append .CreateParameter("optCompletionTime", 134, 1, , Null) '06/09/18 04:44:28 PM
420         .Parameters.Append .CreateParameter("optCompletionDuration", 5, 1, , Null) 'strCompletionDuration) ' 200-advarchar , 1-adParamInput
430         .Parameters.Append .CreateParameter("optClient", 200, 1, 50, strClient) ' Demo Center '''200-advarchar , 1-adParamInput
440         .Parameters.Append .CreateParameter("optLocation", 200, 1, 255, strLocation) '"" '''200-advarchar , 1-adParamInput
450         .Parameters.Append .CreateParameter("optResource", 200, 1, 100, strResource) '"" '''200-advarchar , 1-adParamInput
460         .Parameters.Append .CreateParameter("optModule", 200, 1, 255, "SAM") '"" '''200-advarchar , 1-adParamInput
470         .Parameters.Append .CreateParameter("optAssignTo", 3, 1, , 17) '0 '''200-advarchar , 1-adParamInput  '' ENG_DEVICEMONITORING
480         .Parameters.Append .CreateParameter("OutPutID", 5, 2)    ' 200-advarchar , 1-adParamInput
490         .Parameters.Append .CreateParameter("optAlertGroup", 200, 1, 255, "PMT") '"" '''200-advarchar , 1-adParamInput
500         .Parameters.Append .CreateParameter("optMNDateTime", 134, 1, , Format(Now(), "MM\/DD\/YYYY HH:NN:SS AMPM")) 'now() '''200-advarchar , 1-adParamInput
510         .Parameters.Append .CreateParameter("optThresholdCounter", 200, 1, 2000, strThresholdCounter) '"" '''200-advarchar , 1-adParamInput
520         .Parameters.Append .CreateParameter("optThresholdValue", 200, 1, 255, Null) 'strThresholdValue) '"" '''200-advarchar , 1-adParamInput
530         .Parameters.Append .CreateParameter("optThresholdStatus", 200, 1, 255, strThresholdStatus) '"" '''200-advarchar , 1-adParamInput
540         .Parameters.Append .CreateParameter("optAdditionalDetails", 200, 1, 255, strAdditionalDetails) '"" '''200-advarchar , 1-adParamInput
550         .Parameters.Append .CreateParameter("optTaskID", 5, 1, , Null) 'intTaskID) '0
560         .Parameters.Append .CreateParameter("optTabID", 3, 1, , 3)   '' Others TAB
570         .Parameters.Append .CreateParameter("optVer", 3, 1, , intVer) '0
580         .Parameters.Append .CreateParameter("OptRegID", 5, 1, , intRegID) '-1
            
590         .Parameters.Append .CreateParameter("@optImmResolution", adInteger, adParamInput)
600         .Parameters.Append .CreateParameter("@optUserImpact", adInteger, adParamInput)
610         .Parameters.Append .CreateParameter("@Srno", adVarChar, adParamInput, 100)
620         .Parameters.Append .CreateParameter("@INSuspensionFlag", adVarChar, adParamInput, 1, INSuspensionFlag)
630         If INDATETime <> "12:00:00 AM" Then
640               .Parameters.Append .CreateParameter("@INDATETime", adDate, adParamInput, , INDATETime)
650         Else
660               .Parameters.Append .CreateParameter("@INDATETime", adDate, adParamInput)
670         End If
680         .Parameters.Append .CreateParameter("@ConditionID", adBigInt, adParamInput, , lngCondID)
690         .Execute
700      End With
          
710      If ERR.Number = 0 Then
720         lngJobID = CMDJobMgmt(17)
730         fnCreateNewJob = True
740      Else
750         lngJobID = 0
760         fnCreateNewJob = False
770      End If
780      Set CMDJobMgmt = Nothing
790   Exit Function
bye:
800       LogNotify "fnCreateNewJob", "INFO", ERR.Number, ERR.Description, Erl
End Function
Public Function NewJob(sMSPName As String, sJobName As String, sJobDesc As String, sJobCategoryName As String, sSiteCode As String, sSiteSubCode As String, lngRegID As Long, sResource As String) As Double
      Dim CMDJobMgmt As New ADODB.Command

10    On Error GoTo ErrHandler
            
20       With CMDJobMgmt
30          .ActiveConnection = CnNOCBO
40          .CommandText = "JobManagement_sod"
50          .CommandType = adCmdStoredProc
60          .Parameters.Append .CreateParameter("Return", adDouble, 4)
            
70          .Parameters.Append .CreateParameter("MSPName", adVarChar, adParamInput, 100, sMSPName)
80          .Parameters.Append .CreateParameter("JobName", adVarChar, adParamInput, 255, sJobName)
90          .Parameters.Append .CreateParameter("JobDescription", 203, adParamInput, Len(sJobDesc) + 1, sJobDesc)
100         .Parameters.Append .CreateParameter("JobTypeID", adInteger, adParamInput, , 3)
110         .Parameters.Append .CreateParameter("CategoryName", adVarChar, adParamInput, 5000, sJobCategoryName)
120         .Parameters.Append .CreateParameter("intCatType", adInteger, adParamInput, , 2)
130         .Parameters.Append .CreateParameter("intStatusID", adInteger, adParamInput, , 0)
140         .Parameters.Append .CreateParameter("Ispublish", adInteger, adParamInput, , Null)
150         .Parameters.Append .CreateParameter("StartDateTime", adDBTime, adParamInput, , Format(Now(), "MM\/DD\/YYYY HH:NN:SS AMPM"))
160         .Parameters.Append .CreateParameter("optCompletionTime", adDBTime, adParamInput, , Null)
170         .Parameters.Append .CreateParameter("optCompletionDuration", 5, adParamInput, , Null)
180         .Parameters.Append .CreateParameter("optClient", adVarChar, adParamInput, 50, sSiteCode)
190         .Parameters.Append .CreateParameter("optLocation", adVarChar, adParamInput, 255, sSiteSubCode)
200         .Parameters.Append .CreateParameter("optResource", adVarChar, adParamInput, 100, sResource)
210         .Parameters.Append .CreateParameter("optModule", adVarChar, adParamInput, 255, "SDM")
220         .Parameters.Append .CreateParameter("optAssignTo", adInteger, adParamInput, , 17)
230         .Parameters.Append .CreateParameter("OutPutID", 5, adParamOutput)
240         .Parameters.Append .CreateParameter("optAlertGroup", adVarChar, adParamInput, 255, "PMT")
250         .Parameters.Append .CreateParameter("optMNDateTime", adDBTime, adParamInput, , Format(Now(), "MM\/DD\/YYYY HH:NN:SS AMPM"))
260         .Parameters.Append .CreateParameter("optThresholdCounter", adVarChar, adParamInput, 2000, "")
270         .Parameters.Append .CreateParameter("optThresholdValue", adVarChar, adParamInput, 255, Null)
280         .Parameters.Append .CreateParameter("optThresholdStatus", adVarChar, adParamInput, 255, "")
290         .Parameters.Append .CreateParameter("optAdditionalDetails", adVarChar, adParamInput, 255, "")
300         .Parameters.Append .CreateParameter("optTaskID", 5, adParamInput, , Null)
310         .Parameters.Append .CreateParameter("optTabID", adInteger, adParamInput, , Null)
320         .Parameters.Append .CreateParameter("optVer", adInteger, adParamInput, , 1)
330         .Parameters.Append .CreateParameter("OptRegID", 5, adParamInput, , lngRegID)
            
340         .Execute
350      End With
          
360      NewJob = CMDJobMgmt("OutPutID").Value
         
          
370      Set CMDJobMgmt = Nothing
         
380   Exit Function
ErrHandler:
390       LogNotify "New Job", "INFO", ERR.Number, ERR.Description, Erl
End Function


'--------------------------------------------------------------------------------------------------------------
' This function will add more comments to particular job.
' Return : True : If sucessful, False : If error occured.
' intUpdateOrCloseJob : 0 = Update Job 1 = closed.
'--------------------------------------------------------------------------------------------------------------
Public Function UpdateJob(ByVal lngJobID, ByVal strComments, ByVal intUpdateOrCloseJob, ByVal strClient, ByVal strLocation, ByVal strGroupName)
      Dim cmdJobMgmtupdate As ADODB.Command
10    On Error GoTo Er
          
20        UpdateJob = False
30        If CDbl(lngJobID) <= 0 Then Exit Function
          
40        If intUpdateOrCloseJob > 1 Then Exit Function
50        Set cmdJobMgmtupdate = New ADODB.Command
          
60        With cmdJobMgmtupdate
70            .ActiveConnection = CnNOCBO
80            .CommandText = "UpdateJobForPrevMaint" '"UpdateJob"
90            .CommandType = adCmdStoredProc
100           .Parameters.Append .CreateParameter("Return", adDouble, 4)  'Return parameter
110           .Parameters.Append .CreateParameter("lngJobID", adDouble, 1, , lngJobID)  ' 200-advarchar , 1-adParamInput
120           .Parameters.Append .CreateParameter("strComments", adLongVarWChar, 1, Len(strComments) + 1, strComments) ' 200-advarchar , 1-adParamInput
130           .Parameters.Append .CreateParameter("intUpdateOrCloseJob", adInteger, 1, , intUpdateOrCloseJob) '
140           .Parameters.Append .CreateParameter("Groupname", adVarChar, 1, 250, strGroupName)   ' 200-advarchar , 1-adParamInput
150           .Parameters.Append .CreateParameter("Client", adVarChar, 1, 250, strClient)  ' 200-advarchar , 1-adParamInput
160           .Parameters.Append .CreateParameter("Location", adVarChar, 1, 250, strLocation) ' 200-advarchar , 1-adParamInput
170       End With
          
180       cmdJobMgmtupdate.Execute
190       If cmdJobMgmtupdate(0) = 0 Then
200           UpdateJob = True
210       Else
220           If cmdJobMgmtupdate(0) = -1 Then LogNotify "UpdateJob", "INFO", "Job is already closed"
230           UpdateJob = False
240       End If
          
250       Set cmdJobMgmtupdate = Nothing
          
260   Exit Function
Er:
270       LogNotify "UpdateJob", "INFO", ERR.Number, ERR.Description, Erl
End Function


Public Function DeviceDownExists(Regid As Long) As Boolean
10    On Error GoTo Er
      Dim CmdDeviceDown As New ADODB.Command

20       With CmdDeviceDown
30          .ActiveConnection = CnAlertGen
40          .CommandType = adCmdStoredProc
50          .CommandText = "USP_EnggDeviceDown_Exists"
80          .Parameters.Append .CreateParameter("@inRegid", adBigInt, adParamInput, , Regid)
90          .Parameters.Append .CreateParameter("@OutStatus", adTinyInt, adParamOutput)
110         .Execute

130         If .Parameters("@OutStatus") = 1 Then
140              LogNotify "DeviceDownExists", "INFO", "USP_EnggDeviceDown_Exists Devicedown Alert found"
150              DeviceDownExists = True
160              Exit Function
170         Else
180              LogNotify "DeviceDownExists", "INFO", "USP_EnggDeviceDown_Exists Devicedown Alert not found"
190              DeviceDownExists = False
200              Exit Function
210         End If
220      End With
       
         
230   Exit Function
Er:
240       DeviceDownExists = False
250       LogNotify "DeviceDownExists", "ERROR", ERR.Number, ERR.Description, Erl
End Function


Public Function fnCloseRefTkt(ByVal strAlertId As String, ByVal strModulename As String, ByVal strExeName As String, Optional ByVal intAlertStatus As Integer)
10     On Error GoTo Er
       
       Dim cmdRefTkt As ADODB.Command
20     Set cmdRefTkt = New ADODB.Command
30
40        fnCloseRefTkt = False
50        With cmdRefTkt
60            .ActiveConnection = CnNOCBO
70            .CommandText = "USP_Close_RefTicketId"
80            .CommandType = adCmdStoredProc
90            .Parameters.Append .CreateParameter("AlertId", adVarChar, adParamInput, 100, strAlertId)
100           .Parameters.Append .CreateParameter("Modulename", adVarChar, adParamInput, 100, strModulename)
110           .Parameters.Append .CreateParameter("ExeName", adVarChar, adParamInput, 100, strExeName)  '
120           .Parameters.Append .CreateParameter("AlertStatus", adInteger, adParamInput, , intAlertStatus)
130           .Parameters.Append .CreateParameter("OutPutsuccess", adTinyInt, adParamOutput)
140       End With
150       cmdRefTkt.Execute
160       If cmdRefTkt("OutPutsuccess").Value = 1 Then
170           fnCloseRefTkt = True
180       Else
190           fnCloseRefTkt = False
200       End If

210      Set cmdRefTkt = Nothing

220   Exit Function
Er:
230       fnCloseRefTkt = False
240       Set cmdRefTkt = Nothing
250       LogNotify "fnCloseRefTkt", "ERROR", ERR.Number, ERR.Description, Erl
End Function


