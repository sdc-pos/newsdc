Attribute VB_Name = "SYUDUP"
Option Explicit
'********************************************************************
'*
'*              出荷予定重複データ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const SYUDUP_ID = "SYUDUP"

'ファイル№
Global SYUDUP_No As Integer
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'-------------------------------------------'
'レコード定義
Type SYUDUPREC_Tag
    JGYOBU(0 To 7)              As Byte     '事業場
    DATA_KBN(0 To 0)            As Byte     'データ区分
    TORI_KBN(0 To 1)            As Byte     '取引区分
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '品目番号
    DEN_NO(0 To 9)              As Byte     '伝票番号
    SURYO(0 To 6)               As Byte     '出庫数量
    MUKE_CODE(0 To 7)           As Byte     '得意先コード
    SYUKO_SYUSI(0 To 1)         As Byte     '出庫収支
    SYUKA_YMD(0 To 7)           As Byte     '出荷日付
    ODER_NO(0 To 11)            As Byte     'オーダー番号
    ITEM_NO(0 To 4)             As Byte     'アイテム番号
    MUKE_NAME(0 To 23)          As Byte     '得意先名称
    CHU_KBN(0 To 0)             As Byte     '注文区分
    CHU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
    EXPORT_KBN(0 To 0)          As Byte     '輸出出荷検査区分
    LABEL_ISSUE_KBN(0 To 0)     As Byte     '個装ラベル発行区分
    LABEL_ISSUE_UNIT(0 To 4)    As Byte     '個装ラベル発行単位数
    LABEL_TANKA_KBN(0 To 0)     As Byte     '個装ラベル単価表示区分
    TANKA(0 To 9)               As Byte     '単価
    KINGAKU(0 To 9)             As Byte     '金額
    BIKOU2(0 To 19)             As Byte     '備考２
    REBATE_KBN(0 To 0)          As Byte     'リベート区分
    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
    ATAISA_KBN(0 To 0)          As Byte     '値差区分
    REP_KISHU(0 To 19)          As Byte     '代表機種
    NS__KANRI_NO(0 To 8)        As Byte     'ＮＳ管理番号
    MTS_HIN_CODE(0 To 10)       As Byte     'ＭＴＳ部品コード
    BIKOU1(0 To 39)             As Byte     '備考１
    CHOKU_KBN(0 To 0)           As Byte     '直送区分
    REBATE_RATE(0 To 4)         As Byte     'リベート率
    HIN_NAME(0 To 19)           As Byte     '品名
    JGYOBU_GAI(0 To 7)          As Byte     '対外事業場
    SS_CODE(0 To 7)             As Byte     '直送先コード
    CRLF(0 To 1)                As Byte     'CRLF
End Type

'データ・バッファ
Public SYUDUPREC As SYUDUPREC_Tag
Function SYUDUP_Open() As Integer
'********************************************************************
'*
'*              出荷予定重複ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引数　:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Op_Err     'ｴﾗｰﾄﾗｯﾌﾟON

    SYUDUP_Open = True
                                    
    If GetIni("FILE", SYUDUP_ID, "SYS", c) Then
        Call Log_Out(LOG_F, "SYS.INI [SYUDUP]読み込みエラー")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    SYUDUP_No = FreeFile

    Open FullPath For Binary As #SYUDUP_No
    
    SYUDUP_Open = False

    Exit Function

HS_SIJ_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
                If ans = vbYes Then
                    Resume
                End If
            End If
        Case ErrDeviceUnavailable
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブまたはパスが見つかりません" & FullPath, vbExclamation)
            End If
        Case ErrNotFound
            If Mode = 1 Then
                Beep
                ans = MsgBox("ファイルが見つかりません" & FullPath, vbExclamation)
            End If
        Case Else
            If Mode = 1 Then
                Beep
                ans = MsgBox("エラー [HS_SIJ Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
Function SYUDUP_Get() As Integer
'********************************************************************
'*
'*              出荷予定重複ﾃﾞｰﾀ  ＧＥＴ
'*
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Put_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    SYUDUP_Get = True

    Get #SYUDUP_No, , SYUDUPREC

    SYUDUP_Get = False
    
    Exit Function

SYUDUP_Put_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
            ans = MsgBox("ドライブまたはパスが見つかりません" & SYUDUP_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [SYUDUP Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
End Function
Function SYUDUP_Put(Put_Kbn As Integer) As Integer
'********************************************************************
'*
'*              出荷予定重複ﾃﾞｰﾀ  ＰＵＴ
'*
'*　　　引数　：「０」 保留データへＰＵＴ
'*　　　　　　　「１」 取込データへＰＵＴ
'*
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    On Error GoTo SYUDUP_Put_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    SYUDUP_Put = True

    If Put_Kbn = 0 Then
        Put #SYUDUP_No, , SYUDUPREC
    Else
        Put #XX_SIJ_No, , SYUDUPREC
    End If

    SYUDUP_Put = False
    
    Exit Function

SYUDUP_Put_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
            ans = MsgBox("ドライブまたはパスが見つかりません" & SYUDUP_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [SYUDUP Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
End Function


