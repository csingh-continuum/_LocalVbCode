Attribute VB_Name = "modLog"
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public logFlName As String

Public Function LogNotify(ByVal pInStrModule As String, _
                          ByVal pInStrErrorType As String, _
                          ByVal pInStrValue1 As String, _
                          Optional ByVal pInStrValue2 As String, _
                          Optional ByVal pInStrValue3 As String = "")
10        On Error GoTo CatchErr
          Dim fi As Long, LogString As String
20        fi = FreeFile
30        LogString = Now() & vbTab & pInStrModule & vbTab & pInStrErrorType & vbTab & pInStrValue1 & vbTab & pInStrValue2 & vbTab & pInStrValue3
40        Open logFlName For Append As #fi
50            Print #fi, LogString
60        Close #fi
70        Exit Function
CatchErr:
80        ERR.Clear
End Function

Public Function CreateLogFile() As Boolean
10        flLogFl = App.Path & "\Log\" & App.EXEName & "_" & Format(Now(), "ddmmyyyy") & ".log"
          Dim WFD As WIN32_FIND_DATA
          Dim hSearch As Long
          Dim fi As Long
20        hSearch = FindFirstFile(flLogFl, WFD)
30        If hSearch <> INVALID_HANDLE_VALUE Then
40            Cont = FindClose(hSearch)
50            logFlName = flLogFl
60        Else
70            lgHeadStr = ""
80            fi = FreeFile()
90            Open flLogFl For Append As #fi
100               lgHeadStr = "#Software: SAAZ.ITS.DC." & App.EXEName
110               Print #fi, lgHeadStr
120               lgHeadStr = "#Version: 2.0"
130               Print #fi, lgHeadStr
140               lgHeadStr = "#Date: " & Now()
150               Print #fi, lgHeadStr
160               lgHeadStr = "#Fields: dtime" & vbTab & "module" & vbTab & "type" & vbTab & "val1" & vbTab & "val2" & vbTab & "val3"
170               Print #fi, lgHeadStr
180           Close #fi
190           logFlName = flLogFl
200       End If
End Function
Public Function CreateDir(ByVal DirName As String) As Boolean
10        On Error GoTo ErrHandler
          Dim ret As Long, tempInt As Integer
          Dim DirMain As String
          Dim Security As SECURITY_ATTRIBUTES
20        DirName = QualifyPath(DirName)
30        DirMain = Split(DirName, "\")(0) & "\" & Split(DirName, "\")(1)
40        For tempInt = 2 To UBound(Split(DirName, "\"))
50            If LenB(Trim$(Dir(DirMain, vbDirectory))) = 0 Then
60                 ret = CreateDirectory(DirMain, Security)
70                 If ret = 0 Then
80                     CreateDir = False
90                     Exit Function
100                End If
110           End If
120           DirMain = DirMain & "\" & Split(DirName, "\")(tempInt)
130       Next
140       CreateDir = True
150       Exit Function
ErrHandler:
160       LogNotify "CreateDir", "ERROR", ERR.Number, ERR.Description
170       ERR.Clear
End Function
Public Function QualifyPath(ByVal sPath As String) As String
10        On Error GoTo ErrHandler
          
20        If Right$(sPath, 1) <> "\" Then
30              QualifyPath = sPath & "\"
40        Else: QualifyPath = sPath
50        End If
60        Exit Function
ErrHandler:
70        LogNotify "QualifyPath", "Error", ERR.Description
End Function

Public Function getProfVals(ByVal secName As String, ByVal keyName As String, ByVal flName As String) As String
10        On Error GoTo errHandle
          Dim NC As Long
          Dim ret As String * 8000
          Dim psValue
20        NC = GetPrivateProfileString(secName, keyName, "", ret, Len(ret), flName)
30        If NC = 0 Then
40            GoTo errHandle
50        Else
60            psValue = Left$(ret, InStr(1, ret, vbNullChar) - 1)
70        End If
80        getProfVals = psValue
90        Exit Function
errHandle:
100       getProfVals = ""
End Function





