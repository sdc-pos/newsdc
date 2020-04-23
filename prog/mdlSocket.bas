Attribute VB_Name = "mdlSocket"
Option Explicit

'[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加

'----- 初期化ファイル(ソケット通信用) -----
Public Const SEC_SOCKET         As String = "F110010"               'ソケット通信用設定セクション
Public Const KEY_LOCALPORT      As String = "SocketPort"            'ローカルポート番号

Public Const DEF_LOCALPORT      As Long = 2222                      'デフォルトローカルポート番号

Public Const SEC_LOG            As String = "F110010"               'ログファイル出力用設定セクション
Public Const KEY_LOGWRITE       As String = "LogWrite"              'ログファイル出力フラグ
Public Const KEY_LOGPATH        As String = "LogPath"               'ログファイル保存フォルダ
Public Const KEY_LOGSAVE        As String = "LogSave"               'ログファイル保存期間

Public Const DEF_LOGWRITE       As Boolean = True                   'デフォルトログファイル出力フラグ
Public Const DEF_LOGSAVE        As Integer = 30                     'デフォルトログファイル保存期間

Public Const FNC_LOGSAVECHK     As String = "ログ保存期間チェック処理"
Public Const FNC_DATEMONITOR    As String = "日付更新監視処理"

Public Const FNC_PARENTCONN     As String = "アクセスポイント接続処理"
Public Const FNC_PARENTDISCONN  As String = "アクセスポイント切断処理"
Public Const FNC_RECVDATA       As String = "データ受信処理"
Public Const FNC_RESPDATA       As String = "データ返信処理"
Public Const FNC_SENDDATA       As String = "データ送信処理"
Public Const FNC_RECVMESSAGE    As String = "メッセージ受信処理"
Public Const FNC_FILESEND       As String = "ファイル送信処理"
Public Const FNC_SOCKCLOSE      As String = "ソケット通信切断処理"
Public Const FNC_SOCKCONNECT    As String = "ソケット通信接続処理"
Public Const FNC_SOCKCONNREQ    As String = "ソケット通信接続要求処理"
Public Const FNC_SOCKSEND       As String = "ソケットデータ送信処理"
Public Const FNC_SOCKRECEIVE    As String = "ソケットデータ受信処理"
Public Const FNC_SOCKERROR      As String = "ソケット通信エラー"

Public Const MAX_FSENDDATASIZE  As Integer = 1445                   '最大ファイル送信データサイズ

Public Const P_FILELOAD         As String = "FILELOAD"              'ファイル受信用コマンド

Public Const RESP_OK            As String = "1"                     '正常終了
Public Const RESP_NG            As String = "9"                     'エラー

Public Type SOCKET_CONFIG
    m_IsListen      As Boolean          'ソケット通信開始フラグ(True:リスナー初期化済み)
    m_LocalPort     As Long             'ローカルポート番号
End Type

Public Type LOG_CONFIG
    m_LogWrite      As Boolean          'ログファイル出力フラグ
    m_LogPath       As String           'ログファイル保存フォルダ
    m_LogFName      As String           'ログファイル名
    m_LogSave       As Long             'ログファイル保存期間(単位[日] 0 の場合は削除しない)
End Type

Public gbl_SockCfg  As SOCKET_CONFIG
Public gbl_LogCfg   As LOG_CONFIG

'----- アプリケーション起動日時 -----
Public gbl_StartApp As Date
    
Public Enum LogMsgIcon
    icoDownload = 1
    icoUpload = 2
    icoError = 3
    icoMessage = 4
End Enum

'*******************************************************************************
' ログメッセージ表示
' process   :   指定されたログメッセージをリストに追加する
' input     :   strMsg      メッセージ文字列
'           :   strFunc     処理名
'           :   [intID]     子機ID
'           :   [strIP]     IPアドレス
'           :   [lngIcon]   アイコン種別
' output    :   なし
' return    :   なし
'*******************************************************************************
Public Sub WriteLogMsg(ByVal strMsg As String, ByVal strFunc As String, Optional ByVal intID As Integer = -1, Optional strIP As String = "", Optional lngIcon As LogMsgIcon = icoMessage)
    Static varIcon  As Variant
    Dim strNow      As String
    Dim intFNo      As Integer
    Dim strLogText  As String
    Dim strLogFile  As String
    Dim strID       As String
    Dim strIPAddr   As String
    Dim flgOpen     As Boolean
    
    If IsArray(varIcon) = False Then
        varIcon = Array("(↓)", "(↑)", "(×)", "(ｉ)")
    End If
    
    '----- 現在日時の文字列作成 -----
    strNow = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    
    '----- 子機IDの文字列作成 -----
    strID = IIf(intID < 0, "---", Format$(intID, "000"))
    
    strIPAddr = IIf(Len(strIP) > 0, "(" & strIP & ")", "")
    
    '----- ログ出力文字列を作成 -----
    strLogText = "[" & strNow & "]" & vbTab & CStr(varIcon(lngIcon - 1)) & vbTab & "(" & strID & ")" & vbTab & strIPAddr & vbTab & "【" & strFunc & "】" & vbTab & strMsg
    
    '----- ログ出力オプション -----
    If gbl_LogCfg.m_LogWrite = False Then Exit Sub
    
    '------------------------
    '   ログファイルに出力
    '------------------------
    flgOpen = False
    
    strLogFile = gbl_LogCfg.m_LogFName
    
    On Error GoTo WriteLogMsg_Exit
    intFNo = FreeFile
    Open strLogFile For Append Access Write Lock Read Write As #intFNo
    flgOpen = True
    
    '----- ログメッセージを書きこみ -----
    Print #intFNo, strLogText
    
WriteLogMsg_Exit:
    
    If flgOpen = True Then
        Close #intFNo
    End If

End Sub

'******************************************************************************************
' 実行時エラーメッセージ出力
' process   :   処理中に発生した実行時エラーのメッセージをログに出力する。
' input     :   objErr      エラーオブジェクト
'           :   strFunc     実行処理名
' output    :   なし
' return    :   エラーメッセージ
'******************************************************************************************
Public Sub WriteLogErr(objErr As ErrObject, strFunc As String)
    Dim strMsg  As String
    Dim lngIcon As Long
    
    '----- メッセージの作成とアイコンの決定 -----
    If objErr.Number = 0 Then
        strMsg = "エラーは発生していません。"
        lngIcon = LogMsgIcon.icoMessage
    Else
        strMsg = strFunc & "中に実行時エラーが発生しました。[内容] : " & ChStr(objErr.Description, vbCrLf, " ") & " [番号] : " & CStr(objErr.Number)
        lngIcon = LogMsgIcon.icoError
    End If
    
    '----- ログ出力 -----
    Call WriteLogMsg(strMsg, strFunc, , , lngIcon)
End Sub

'*******************************************************************************
' ログファイル削除
' process   :   一定期間以上前のログファイルを削除する
' input     :   strTagName  ログファイル名称
'           :   lngSaveDay  ログ保持期間
' output    :   なし
' return    :   [strTagName]_yyyymmdd.log の形式のファイル名を検索します。
'*******************************************************************************
Public Sub DeleteLogFile(strTagName As String, lngSaveDay As Long)
    Dim strPath As String
    Dim strFind As String
    Dim strSave As String
    Dim strCheck As String
    Dim strFName As String
    
    Call WriteLogMsg(FNC_LOGSAVECHK & "を開始します。(検索対象:" & strTagName & "_YYYYMMDD.log / 保存期間:" & CStr(lngSaveDay) & " 日間)", FNC_LOGSAVECHK, , , icoMessage)
    
    On Error GoTo DeleteLogFile_Exit
    Err.Clear
    
    If lngSaveDay <= 0 Then Exit Sub
    
    '----- 検索用文字列作成 -----
    strFind = GetFullPath(gbl_LogCfg.m_LogPath, strTagName) & "_????????.log"
    
    '----- 保存期間の日付文字列を作成 -----
    strSave = Format$(DateAdd("d", -(lngSaveDay), gbl_StartApp), "yyyymmdd")
    
    '----- ログファイル検索開始 -----
    strCheck = Dir$(strFind, vbArchive)
    Do While Len(strCheck) > 0
        If StrComp(Mid$(strCheck, Len(strTagName) + 2, 8), strSave, vbTextCompare) < 0 Then
            '----- 保存期間より古いログファイルの場合 -----
            strFName = GetFullPath(gbl_LogCfg.m_LogPath, strCheck)
            Kill strFName
            Call WriteLogMsg("保存期間より古いログファイル(" & strFName & ")を削除しました。", FNC_LOGSAVECHK, , , icoMessage)
        End If
        
        '----- 次のファイルを検索 -----
        strCheck = Dir$
    Loop
    
DeleteLogFile_Exit:
    If Err.Number <> 0 Then
        Call WriteLogErr(Err, FNC_LOGSAVECHK)
    End If
    
    Call WriteLogMsg(FNC_LOGSAVECHK & "を終了します。", FNC_LOGSAVECHK, , , icoMessage)
End Sub

'*******************************************************************************
' 送信用データ加工
' process   :   指定された文字列をハンディへの送信用データに加工する
' input     :   strMsg      送信データ
' output    :   なし
' return    :   送信用文字列
'*******************************************************************************
Public Function ConvTextMsg(strMsg As String) As String
    Dim strBuf As String
    Dim intLoop As Integer
    
    '----- 送信文字列を加工 -----
    strBuf = ChStr(strMsg, "[CRLF]", vbCrLf)
    strBuf = ChStr(strBuf, "[CR]", vbCr)
    strBuf = ChStr(strBuf, "[LF]", vbLf)
    strBuf = ChStr(strBuf, "[TAB]", vbTab)
    strBuf = ChStr(strBuf, "[ESC]", Chr$(27))
    For intLoop = 1 To 255
        strBuf = ChStr(strBuf, "[0x" & Right$("0" & UCase(Hex(intLoop)), 2) & "]", Chr$(intLoop))
        strBuf = ChStr(strBuf, "[0x" & Right$("0" & LCase(Hex(intLoop)), 2) & "]", Chr$(intLoop))
    Next intLoop
    strBuf = ChStr(strBuf, "[DATE]", Format$(Now, "yyyymmdd"))
    strBuf = ChStr(strBuf, "[TIME]", Format$(Now, "hhmmss"))
    
    ConvTextMsg = strBuf
End Function

'*******************************************************************************
' 受信データ加工
' process   :   ハンディから受信したデータをテキスト形式に加工する
' input     :   strMsg      受信データ
' output    :   なし
' return    :   受信テキスト文字列
'*******************************************************************************
Public Function ConvBinaryMsg(strMsg As String) As String
    Dim strBuf As String
    Dim strWrk As String
    Dim strChr As String
    Dim intLoop As Integer
    
    '----- 受信文字列を加工 -----
    strBuf = ChStr(strMsg, vbCrLf, "[CRLF]")
    strBuf = ChStr(strBuf, vbCr, "[CR]")
    strBuf = ChStr(strBuf, vbLf, "[LF]")
    strBuf = ChStr(strBuf, vbTab, "[TAB]")
    strBuf = ChStr(strBuf, Chr$(27), "[ESC]")
    strWrk = ""
    For intLoop = 1 To Len(strBuf)
        strChr = Mid$(strBuf, intLoop, 1)
        If (Asc(strChr) < 0) Or (Asc(strChr) >= &H20 And Asc(strChr) <= &H7E) Or (Asc(strChr) >= &HA0 And Asc(strChr) <= &HDF) Then
            '----- 表示可能文字の場合 -----
            strWrk = strWrk & strChr
        Else
            '----- 表示不可能文字の場合 -----
            strWrk = strWrk & "[0x" & Right$("0" & UCase(Hex(Asc(strChr))), 2) & "]"
        End If
    Next intLoop
    
    ConvBinaryMsg = strWrk
End Function

'********************************************************************************
' フルパスファイル名作成
' process   :   ディレクトリ名、ファイル名を元にフルパスのファイル名を作成する
' intput    :   strPath     パス名
'           :   strFile     ファイル名
'           :   varToken    パス名、ファイル名の区切り文字
' output    :   なし
' return    :   フルパスファイル名
'********************************************************************************
Public Function GetFullPath(strPath As String, strFile As String, Optional varToken As Variant = "\") As String
    Dim strFPath     As String
    Dim strName1    As String
    Dim strName2    As String
    Dim strMake     As String
    Dim intTokLen   As Integer
    
    If strPath = "" And strFile <> "" Then
        'ファイルだけが指定されている
        strFPath = strFile
        GoTo GetFullPath_Exit
    ElseIf strPath <> "" And strFile = "" Then
        'パスだけが指定されている
        strFPath = strPath
        GoTo GetFullPath_Exit
    ElseIf strPath = "" And strFile = "" Then
        'どちらも指定されていない
        strFPath = ""
        GoTo GetFullPath_Exit
    End If
    
    intTokLen = Len(CStr(varToken))
    
    If Right$(strPath, intTokLen) = CStr(varToken) Then
        '区切り文字を省く
        strName1 = Left$(strPath, Len(strPath) - intTokLen)
    Else
        strName1 = strPath
    End If
    If Left$(strFile, intTokLen) = CStr(varToken) Then
        '区切り文字を省く
        strName2 = Right$(strFile, Len(strFile) - intTokLen)
    Else
        strName2 = strFile
    End If
    
    'フルパスを作成
    strFPath = strName1 & CStr(varToken) & strName2
    
GetFullPath_Exit:
    GetFullPath = strFPath
End Function

'******************************************************************************************
' ディレクトリ作成
' process   :   指定されたパスを作成する。
' input     :   strPath     作成するパス
' output    :   無し
' date      :   1999/02/25 - K.HAYASHI 修正
' return    :   True:正常終了, False:異常終了
'******************************************************************************************
Public Function MkDirEx(strPath As String) As Boolean
    Dim strDrv   As String
    Dim flgRet   As Variant
    Dim strDir() As String
    Dim strWork  As String
    Dim intCnt   As Integer
    Dim intLoop  As Integer
    
    Err.Clear
    On Error GoTo MkDirEx_Exit
    
    '----- リターン値初期化 -----
    flgRet = False
    
    '----- ドライブチェック -----
    If InStr(strPath, ":\") > 0 Then
        '----- ドライブ文字抜き出し -----
        strDrv = Left$(strPath, 2)
        
        Call ChDir(strDrv)
    End If
    
    '----- 各ディレクトリを分割 -----
    intCnt = Explode(Mid$(strPath, 4), strDir, "\")
    If intCnt > 0 Then
        strWork = strDrv
        intLoop = 0
        For intLoop = 0 To intCnt - 1
            If Len(strDir(intLoop)) = 0 Then
                Exit For
            End If
            
            '----- ディレクトリ名作成 -----
            strWork = strWork & "\" & strDir(intLoop)
            
            '----- ディレクトリ有無チェック -----
            If Len(Dir2(strWork, vbDirectory)) = 0 Then
                '----- ディレクトリ作成 -----
                MkDir strWork
            End If
        Next intLoop
        
        flgRet = True
    End If
MkDirEx_Exit:
    MkDirEx = flgRet
End Function

'*******************************************************************************
' リスト文字列取得
' process   :   1つにまとめられたリスト文字列を配列に格納する
' input     :   strList     リスト文字列
'           :   varBuf      文字列格納バッファ
'           :   strToken    文字列を区切るトークン
' output    :   varBuf      配列に格納されたリスト文字列
' return    :   配列数, 0:データ無
'*******************************************************************************
Public Function Explode(strList As String, varBuf As Variant, strToken As String) As Integer
    Dim Index As Integer
    Dim intSp As Integer
    Dim intEp As Integer
    
    '----- リターン値初期化 -----
    Index = 0
    
    '----- 引数チェック -----
    If strToken = "" Then
        GoTo Explode_Exit
    End If
    
    ReDim varBuf(0) As String
    intSp = 1
    intEp = InStr(intSp, strList, strToken)
    Do While intEp > 0
        ReDim Preserve varBuf(Index) As String
        varBuf(Index) = Mid$(strList, intSp, intEp - intSp)
        intSp = intEp + Len(strToken)
        intEp = InStr(intSp, strList, strToken)
        Index = Index + 1
    Loop
    ReDim Preserve varBuf(Index) As String
    varBuf(Index) = Mid$(strList, intSp)
    Index = Index + 1
    
Explode_Exit:
    Explode = Index
End Function

'*******************************************************************************
' ファイル有無チェック
' process   :   指定ファイル名が存在するかどうかチェックする
' input     :   szFilePath      確認するファイルのパス
'           :   intAttr           確認するファイルのタイプ
' output    :   なし
' return    :   有:ファイル名, 無:""
'*******************************************************************************
Public Function Dir2(Optional varFilePath As Variant, Optional intAttr As VbFileAttribute = vbNormal) As String
    Static strPath As String
    Dim strBuf As String
    Dim strRet As String
    Dim strChk  As String
    
    On Error Resume Next
    
    If IsMissing(varFilePath) = False Then
        strPath = ""
        Call DivFileName(CStr(varFilePath), strPath, strBuf)
    End If
    
    strRet = Dir$(varFilePath, intAttr)
    If Len(strRet) > 0 Then
        If Len(strPath) > 0 Then
            strChk = GetFullPath(strPath, strRet)
        Else
            strChk = strRet
        End If
        
        If (GetAttr(strChk) And intAttr) <> intAttr Then
            'ファイル属性チェック
            strRet = ""
        End If
    End If
    
    On Error GoTo 0
    
    Dir2 = strRet
End Function

'*******************************************************************************
' ファイル名分解
' process   :   指定ファイル名のディレクトリとファイルの部分を分解する
' input     :   strFilePath     分解するファイルのパス
'           :   strPathName     分解後文字列格納バッファ１
'           :   strFileName     分解後文字列格納バッファ２
'           :   strToken        ファイル名区切り文字
' output    :   strPathName     分解後文字列１
'           :   strFileName     分解後文字列２
' return    :   True:正常, False:異常
'*******************************************************************************
Public Function DivFileName(strFilePath As String, strPathName As String, strFileName As String, Optional strToken As String = "\") As Boolean
    Dim strFBuf As String
    Dim strPath As String
    Dim strFile As String
    Dim intPos  As Integer
    
    If StrComp(Right$(strFilePath, Len(strToken)), strToken) = 0 Then
        '----- 区切り文字が最後にある場合は省く -----
        strFBuf = Left$(strFilePath, Len(strFilePath) - Len(strToken))
    Else
        strFBuf = strFilePath
    End If
    
    For intPos = Len(strFBuf) To 1 Step -1
        '----------------------------------------
        '   最後の文字から区切り文字を検索して
        '   最初に該当した部分以降がファイル名
        '----------------------------------------
        If StrComp(Mid$(strFBuf, intPos, Len(strToken)), strToken) = 0 Then
            strFile = Mid$(strFBuf, intPos + Len(strToken))
            strPath = Left$(strFBuf, intPos - 1)
            If StrComp(strToken, "/") = 0 Then
                If Len(strPath) = 0 Then
                    strPath = "/"
                End If
            ElseIf StrComp(strToken, "\") = 0 Then
                If StrComp(Right$(strPath, 1), ":") = 0 Then
                    '----- ルートパス(C:\等)のとき -----
                    strPath = strPath & strToken
                End If
            End If
            Exit For
        End If
    Next intPos
    
    If Len(strFile) * Len(strPath) <> 0 Then
        strPathName = strPath
        strFileName = strFile
        DivFileName = True
    Else
        strPathName = ""
        strFileName = ""
        DivFileName = False
    End If
End Function

'******************************************************************************************
' 文字列置き換え
' process   :   指定された文字列の中の指定文字を全て置き換える
' input     :   strText     編集文字列
'           :   strToken1   検索文字列
'           :   strToken2   置き換え文字列
' output    :   無し
' return    :   編集後の文字列
'******************************************************************************************
Public Function ChStr(strText As String, strToken1 As String, strToken2 As String) As String
    Dim strWork As String
    Dim strLeft As String
    Dim strRight As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    strWork = strText
    intPos1 = 1
    Do
        '置き換え文字列を検索
        intPos2 = InStr(intPos1, strWork, strToken1)
        If intPos2 = 0 Then
            Exit Do
        Else
            strLeft = Left$(strWork, intPos2 - 1)
            strRight = Mid$(strWork, intPos2 + Len(strToken1))
        End If
        
        '文字列を置き換える
        strWork = strLeft & strToken2 & strRight
        
        intPos1 = Len(strLeft) + Len(strToken2) + 1
    Loop
    ChStr = strWork
End Function

'******************************************************************************************
' 文字列バイト調整
' process   :   指定されたバイト数に調整した文字列を作成する。
' input     :   strText     調整する文字列
'           :   intLength   調整するバイト数
'           :   lngFormat   テキスト配置位置
' output    :   なし
' return    :   指定バイト数に調整した文字列
'******************************************************************************************
Function AlignText(strText As String, intLength As Integer, Optional lngFormat As AlignmentConstants = -1) As String
    Dim strMake     As String       '加工文字列バッファ
    Dim strBuf      As String       'ワークバッファ
    Dim strBody     As String       'ワークバッファ
    Dim strLast     As String       'ワークバッファ
    Dim intSpcMax   As Integer      'スペース追加数
    Dim intSpcL     As Integer      '左側スペース数
    Dim intSpcR     As Integer      '右側スペース数
    
    If Len(strText) = 0 Then
        '----- 長さ0の文字の場合 -----
        AlignText = Space$(intLength)
        Exit Function
    End If
    
    '----- 指定されたバイト数で文字列を切る -----
    strBuf = StrConv(LeftB(StrConv(strText, vbFromUnicode), intLength), vbUnicode)
    
    '----- 1文字分前の文字列と最後の1文字を取得 -----
    strBody = Left$(strText, Len(strBuf) - 1)
    strLast = Mid$(strText, Len(strBuf), 1)
    
    '----- 最後の文字のバイト数をチェック -----
    If LenB(StrConv(strLast, vbFromUnicode)) = 2 Then
        '----- 2バイト文字の場合は合計バイト数をチェック -----
        If LenB(StrConv(strBody, vbFromUnicode)) + LenB(StrConv(strLast, vbFromUnicode)) > intLength Then
            '----- 指定バイト数より大きい場合は1文字分省く -----
            strMake = strBody
        Else
            '----- そのまま結合する -----
            strMake = strBody & strLast
        End If
    Else
        '----- 最後の文字が1バイト文字ということは正常なサイズ -----
        strMake = strBody & strLast
    End If
    
    '----- 文字位置設定 -----
    intSpcMax = intLength - LenB(StrConv(strMake, vbFromUnicode))
    
    Select Case (lngFormat)
    Case vbCenter
        '----- 中央揃え -----
        intSpcL = intSpcMax \ 2
        intSpcR = intSpcMax - intSpcL
        AlignText = Space$(intSpcL) & strMake & Space$(intSpcR)
    Case vbRightJustify
        '----- 右詰め -----
        AlignText = Space$(intSpcMax) & strMake
    Case vbLeftJustify
        '----- 左詰め -----
        AlignText = strMake & Space$(intSpcMax)
    Case Else
        '----- 設定なし -----
        AlignText = strMake
    End Select
    
End Function

Public Function GetFileData(ByVal FileName As String) As String
    Dim strRetBuf  As String
    Dim strReadBuf  As String
    Dim intFNo      As Integer
    Dim blnOpen     As Boolean
    
    On Error GoTo GetFileData_Exit
    
    '----- 戻り値の初期化 -----
    strRetBuf = ""
    
    blnOpen = False
    
'    If IsExist(FileName) = False Then
'        '----- ファイルが存在しない場合 -----
'        GoTo GetFileData_Exit
'    End If
    
    '----------------------
    '   ファイルオープン
    '----------------------
    intFNo = FreeFile()
    Open FileName For Input Access Read Lock Write As #intFNo
    blnOpen = True
    
    Do While Not EOF(intFNo)
        Line Input #intFNo, strReadBuf
        strRetBuf = strRetBuf & strReadBuf & vbCrLf
    Loop
    
GetFileData_Exit:
    If blnOpen = True Then
        '----------------------
        '   ファイルクローズ
        '----------------------
        Close #intFNo
    End If
    
    GetFileData = strRetBuf
End Function

