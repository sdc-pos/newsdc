Attribute VB_Name = "MainF105035"
Option Explicit

'Public Const Last_Update_Day$ = "[F105035] 2017.04.28 14:00"
Public Const Last_Update_Day$ = "[F105035] 2018.10.02 13:10"


Sub Main()
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    Last_JGYOBU = Trim(Command)



    ' 2�d�N���̏ꍇ�́A��O�Ɏ����Ă��Ď������g�͏I������
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle




    F1050351.Show
End Sub
