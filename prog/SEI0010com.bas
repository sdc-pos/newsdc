Attribute VB_Name = "SEI0010com"
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
     ByVal crColor As Long, ByVal wFillType As Long) As Long


Sub MAIN()
    
    
    
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










    SEI00101.Show
End Sub

