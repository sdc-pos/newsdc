Attribute VB_Name = "CHGH"
Option Explicit
'********************************************************************
'*                                                                  *
'*              外部品番変更  ファイル定義                            *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'ファイルＩＤ
Global Const CHGH_ID = "CHGH"

'ファイル№
Global CHGH_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type CHGHREC_Tag
    TEXT_NO(0 To 8) As Byte         'ﾃｷｽﾄ№
    JGYOBU(0 To 0) As Byte          '事業部区分
    CYOK_KBN(0 To 0) As Byte        '直送区分
    DEN_DT(0 To 7) As Byte          '伝票日付
    IO_KBN(0 To 0) As Byte          '入出庫区分
    PM_KBN(0 To 0) As Byte          '赤黒区分
    DEN_SYU(0 To 0) As Byte         '伝票種別
    DEN_NO(0 To 5) As Byte          '伝票№
    CYU_KBN(0 To 0) As Byte         '注文区分
    HIN_GAI(0 To 12) As Byte        '品番（外部）
    HIN_NAI(0 To 12) As Byte        '品番（内部）
    HIN_NAME(0 To 24) As Byte       '品名
    YOTEI_QTY(0 To 5) As Byte       '数量
    YOSAN_FROM(0 To 4) As Byte      '予算単位（元）
    YOSAN_TO(0 To 4) As Byte        '予算単位（先）
    HOST_SOKO(0 To 1) As Byte       '倉庫区分（ﾎｽﾄ）
    HOST_TANA(0 To 7) As Byte       '棚番（ﾎｽﾄ）
    SYUK_CODE(0 To 4) As Byte       '支給先／出荷先
    SYUK_NAME(0 To 19) As Byte      '支給先／出荷先名
    REC_END(0 To 0) As Byte         'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    CR(0 To 0) As Byte              'ｷｬﾘｯｼﾞﾘﾀｰﾝ
    LF(0 To 0) As Byte              'ﾗｲﾝﾌｨｰﾄﾞ
End Type

'データ・バッファ
Global CHGHREC As CHGHREC_Tag
Function CHGH_Open() As Integer
'********************************************************************
'*                                                                  *
'*              外部品番変更保留ﾃﾞｰﾀ  ＯＰＥＮ                        *
'*                                                                  *
'*      戻り値:false 正常                                            *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Op_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    CHGH_Open = False
                            '外部品番変更保留ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", CHGH_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        CHGH_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)

    CHGH_No = FreeFile

    Open FullPath For Binary As #CHGH_No

    Exit Function

CHGH_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo _
                            + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブまたはパスが見つかりません" & FullPath, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("ファイルが見つかりません" & FullPath, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [CHGH Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
    CHGH_Open = True
    Exit Function
End Function
Function CHGH_Get() As Integer
'********************************************************************
'*                                                                  *
'*              外部品番変更保留ﾃﾞｰﾀ  ＧＥＴ　                        *
'*                                                                  *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Get_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    CHGH_Get = False

    Get #CHGH_No, , CHGHREC

    Exit Function

CHGH_Get_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68
    Select Case Err.Number
        Case ErrDiskNotReady        'ﾄﾞﾗｲﾌﾞが正しく準備されていない
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo _
                  + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable   'ﾄﾞﾗｲﾌﾞorﾊﾟｽが見つからない
            Beep
            ans = MsgBox("ドライブまたはパスが見つかりません" & CHGH_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [CHGH Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    CHGH_Get = True
End Function
Function CHGH_Put(Put_Kbn As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              外部品番変更保留ﾃﾞｰﾀ  ＰＵＴ　                        *
'*                                                                  *
'*　　　引数　：「０」 保留データへＰＵＴ                              *
'*　　　　　　　「１」 取込データへＰＵＴ                              *
'*                                                                  *
'*      戻り値:false 正常                                            *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.05.28  S.Shibano                            *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo CHGH_Put_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    CHGH_Put = False

    If Put_Kbn = 0 Then
        Put #CHGH_No, , CHGHREC
    Else
        Put #XX_SIJ_No, , CHGHREC
    End If

    Exit Function

CHGH_Put_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68
    Select Case Err.Number
        Case ErrDiskNotReady        'ﾄﾞﾗｲﾌﾞが正しく準備されていない
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo _
                  + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable   'ﾄﾞﾗｲﾌﾞorﾊﾟｽが見つからない
            Beep
            ans = MsgBox("ドライブまたはパスが見つかりません" & CHGH_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [CHGH Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    CHGH_Put = True
    Exit Function
End Function


