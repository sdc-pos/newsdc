Attribute VB_Name = "Y_SYU_LOG"
Option Explicit
'********************************************************************
'*
'*              出荷予定データログ
'*
'*          CREATE 2001.05.09
'********************************************************************
Public SYUKA_LOGF   As String   '出荷ログファイル名称
Public SYUKA_LOG_ON As Boolean  '出荷ログ出力ＯＮ／ＯＦＦ

Public Function SYUKA_LOGF_GET_PROC() As Integer
'****************************************************
'*      出荷ログファイル名称の取込み
'*
'*  引数 :  なし
'*  戻り値: false       正常
'*          SYS_ERR     継続できない異常
'****************************************************
Dim c       As String
Dim Ret     As Integer

    SYUKA_LOGF_GET_PROC = SYS_ERR
    
                                        '出荷ログ出力有無取り込み
    If GetIni("SYUKA_LOG", StrConv(App.EXEName, vbUpperCase), "SYS", c) Then
        SYUKA_LOG_ON = False
        SYUKA_LOGF_GET_PROC = False
        Exit Function
    End If
    If Trim(c) = "0" Then
        SYUKA_LOG_ON = False
        SYUKA_LOGF_GET_PROC = False
        Exit Function
    End If

    SYUKA_LOG_ON = True

    If GetIni("FILE", "SYU_LOG", "SYS", c) Then
        Exit Function
    End If

    Ret = InStr(1, Trim(c), ".") - 1
    SYUKA_LOGF = Left(Trim(c), Ret) & Right(Format(Date, "yyyymmdd"), 2) & Right(Trim(c), Len(Trim(c)) - Ret)
    SYUKA_LOGF_GET_PROC = False
End Function


Public Sub SYUKA_LOG_OUT_PROC(YOIN1 As String, YOIN2 As String)
'****************************************************
'*      出荷ログファイルの出力
'*
'*  引数 :  出力要因
'*          データ状況
'*  戻り値:  なし
'*  呼び元で保持している最終出荷予定内容を出力する
'*
'****************************************************
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

    
    stream = FreeFile
    
    On Error Resume Next
    
    If Format(Date, "yyyymmdd") <> Format(FileDateTime(SYUKA_LOGF), "yyyymmdd") Then
        Open SYUKA_LOGF For Output As stream
    Else
        Open SYUKA_LOGF For Append As stream
    End If
    prog = StrConv(App.EXEName, vbUpperCase)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Buf = (Buf & " " & YOIN1 & " " & YOIN2 & " ")
    Buf = (Buf & "伝日付：" & StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) & " ")
    Buf = (Buf & "伝ID：" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & " ")
    Buf = (Buf & "伝№：" & StrConv(Y_SYUREC.DEN_NO, vbUnicode) & " ")
    Buf = (Buf & "注区：" & StrConv(Y_SYUREC.CYU_KBN, vbUnicode) & " ")
    Buf = (Buf & "向け先：" & StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & " ")
    Buf = (Buf & "品番：" & StrConv(Y_SYUREC.HIN_NO, vbUnicode) & " ")
    Buf = (Buf & "数：" & Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0"))
    Print #stream, Buf
    Close stream

End Sub

