Attribute VB_Name = "AlertSusp"
'Public CMDSuspStatus As ADODB.Command
'Public CmdEscMail As ADODB.Command
'Public CmdEscMail_IU As ADODB.Command
Option Explicit

Public Const CONDITIONID = 10407
Public SuspToDate As Date
Public Suspended As String

Public Function GetSuspStatus(Regid As Long, CONDITIONID As Long, Filter As String, FilterDelim As String, ByRef ToDate As Date) As String
10    On Error GoTo Er
          Dim CmdSuspStatus As ADODB.Command
          
20        Set CmdSuspStatus = New ADODB.Command
30       With CmdSuspStatus
         
40          .ActiveConnection = CnAlertConfig ''changed from alertsync to alertconfig 02/01/2013
50          .CommandType = adCmdStoredProc
60          .CommandText = "USP_Tkt_Suspenssion_Generic"
            
           
70          .Parameters.Append .CreateParameter("@INRegid", adBigInt, adParamInput, , Regid)
80          .Parameters.Append .CreateParameter("@INConditionID", adBigInt, adParamInput, , CONDITIONID)
90          .Parameters.Append .CreateParameter("@INFilter", adVarChar, adParamInput, 8000, Filter)
100         .Parameters.Append .CreateParameter("@INFilterDelim", adVarChar, adParamInput, 5, FilterDelim)
            
110         .Parameters.Append .CreateParameter("@OUTTodat", adDate, adParamOutput)
120         .Parameters.Append .CreateParameter("@OUTIsSuspenssion", adInteger, adParamOutput)
            
            ''@OUTIsSuspenssion = 1 if suspended
            ''@OUTIsSuspenssion = 0 if NotSuspended
            
130         .Prepared = True
            
140         LogNotify "GetSuspStatus", "INFO", "Executing SP : USP_Tkt_Suspenssion_Generic  for RegId " & Regid
            
150         .Execute
            
160         LogNotify "GetSuspStatus", "INFO", "Output Param OUTIsSuspenssion :  " & .Parameters("@OUTIsSuspenssion")
170         LogNotify "GetSuspStatus", "INFO", "Output Param @OUTTodat :  " & .Parameters("@OUTTodat") & ""
            
180          GetSuspStatus = IIf(.Parameters("@OUTIsSuspenssion") = 1, "Y", "N")
190          If IsNull(.Parameters("@OUTTodat")) = False Then ToDate = .Parameters("@OUTTodat")
200      End With
         
210      Set CmdSuspStatus = Nothing
         
220   Exit Function
Er:
230       LogNotify "GetSuspStatus", "ERROR", ERR.Number, ERR.Description, Erl
End Function

Public Function GetEscMail(Regid As Long, CONDITIONID As Long, Filter As String, FilterDelim As String) As String
10    On Error GoTo Er

          Dim CmdEscMail As ADODB.Command

20       Set CmdEscMail = New ADODB.Command
30       With CmdEscMail
         
40          .ActiveConnection = CnAlertSync
50          .CommandType = adCmdStoredProc
60          .CommandText = "USP_Tkt_EscMailID_Generic"
            

70          .Parameters.Append .CreateParameter("@INRegid", adBigInt, adParamInput, , Regid)
80          .Parameters.Append .CreateParameter("@INConditionID", adBigInt, adParamInput, , CONDITIONID)
90          .Parameters.Append .CreateParameter("@INFilter", adVarChar, adParamInput, 8000, Filter)
100         .Parameters.Append .CreateParameter("@INFilterDelim", adVarChar, adParamInput, 5, FilterDelim)
            
110         .Parameters.Append .CreateParameter("@OUTEMAILID", adVarChar, adParamOutput, 8000)

120         .Prepared = True
                 
130         LogNotify "GetEscMail", "INFO", "Executing SP : USP_Tkt_EscMailID_Generic"
             
140         .Execute
             
150         GetEscMail = .Parameters("@OUTEMAILID") & ""
             
160      End With
         
170     Set CmdEscMail = Nothing
         
180   Exit Function
Er:
190       LogNotify "GetEscMail", "INFO", ERR.Number, ERR.Description, Erl
End Function

Public Function EscMail_IU(Regid As Long, CONDITIONID As Long, Filter As String, Delim As String, TktId As String _
            , SusDateTime As Date, Action As String, MailSubject As String, MailBody As String, _
            MailFrom As String, MailType As Integer, RawDataId As String) As Integer
                  
10    On Error GoTo Er

         Dim CmdEscMail_IU As ADODB.Command
         
20       Set CmdEscMail_IU = New ADODB.Command
30       With CmdEscMail_IU
         
40          .ActiveConnection = CnAlertSync
50          .CommandType = adCmdStoredProc
60          .CommandText = "USP_MstAlertEmail_Generic_IU"
            
       
70          .Parameters.Append .CreateParameter("@INRegid", adBigInt, adParamInput, , Regid)
80          .Parameters.Append .CreateParameter("@INConditionID", adBigInt, adParamInput, , CONDITIONID)
            
90          If Filter = "" Then
100               .Parameters.Append .CreateParameter("@INFilterParam", adVarChar, adParamInput, 2000)
110         Else
120               .Parameters.Append .CreateParameter("@INFilterParam", adVarChar, adParamInput, 2000, Filter)
130         End If
      ''      .Parameters.Append .CreateParameter("@INTKTID", adBigInt, adParamInput, , TktId)
      ''      .Parameters.Append .CreateParameter("@INSUSDTime", adDate, adParamInput, , SusDateTime)
140          If Val(TktId) = 0 Then
150               .Parameters.Append .CreateParameter("@INTKTID", adBigInt, adParamInput, , Null)
160          Else
170               .Parameters.Append .CreateParameter("@INTKTID", adBigInt, adParamInput, , TktId)
180          End If
190         If SusDateTime <> "12:00:00 AM" Then
200             .Parameters.Append .CreateParameter("@INSUSDTime", adDate, adParamInput, , SusDateTime)
210         Else
220            .Parameters.Append .CreateParameter("@INSUSDTime", adDate, adParamInput)
230         End If
240         .Parameters.Append .CreateParameter("@INAction", adVarChar, adParamInput, 1, Action)
250         .Parameters.Append .CreateParameter("@INMailSubject", adVarChar, adParamInput, Len(MailSubject) + 1, MailSubject)
260         .Parameters.Append .CreateParameter("@INMailBody", adVarChar, adParamInput, Len(MailBody) + 1, MailBody)
270         .Parameters.Append .CreateParameter("@INMailfrom", adVarChar, adParamInput, 300, MailFrom)
280         If Delim = "" Then
290               .Parameters.Append .CreateParameter("@INDelimeter", adVarChar, adParamInput, 5)
300         Else
310               .Parameters.Append .CreateParameter("@INDelimeter", adVarChar, adParamInput, 5, Delim)
320         End If
330         .Parameters.Append .CreateParameter("@INMailType", adInteger, adParamInput, , MailType)
340         If RawDataId = "" Then
350               .Parameters.Append .CreateParameter("@INTaskID", adVarChar, adParamInput, 300)
360         Else
370               .Parameters.Append .CreateParameter("@INTaskID", adVarChar, adParamInput, 300, RawDataId)
380         End If
390         .Parameters.Append .CreateParameter("@OUTStatus", adInteger, adParamOutput)
        
            
400         .Prepared = True
            
410          LogNotify "EscMail_IU", "INFO", "Executing SP : USP_MstAlertEmail_Generic_IU"
             
420         .Execute
             
430          LogNotify "EscMail_IU", "INFO", "Output Parameter @OUTStatus : " & .Parameters("@OUTStatus")
              
440          EscMail_IU = .Parameters("@OUTStatus")
450      End With
         
460      Set CmdEscMail_IU = Nothing
470   Exit Function
Er:
480       LogNotify "EscMail_IU", "ERROR", ERR.Number, ERR.Description, Erl
End Function






