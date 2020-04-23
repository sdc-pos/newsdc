Attribute VB_Name = "HS_ZAI1"
Option Explicit
'********************************************************************
'*
'*              在庫設定データ ファイル定義
'*
'*          CREATE 2001.05.18
'********************************************************************
'ファイルＩＤ
Global Const HS_ZAI_ID1 = "HS_ZAI1"         '洗濯機事業部
'ファイル№
Global HS_ZAI_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type HS_ZAIREC_Tag
    JGYOBU(0 To 0) As Byte          '事業部区分
    HOST_SOKO(0 To 1) As Byte       '倉庫区分（ﾎｽﾄ）
    HIN_GAI(0 To 12) As Byte        '品番（外部）
    HIN_NAI(0 To 12) As Byte        '品番（内部）
    HIN_NAME(0 To 24) As Byte       '品名
    HOST_TANA(0 To 7) As Byte       '棚番（ﾎｽﾄ）
    QTY_SIGN(0 To 0) As Byte        '数量サイン
    ZEN_Z_QTY(0 To 6) As Byte       '前日在庫数
    FILLER(0 To 8) As Byte          'FILLER
    REC_END(0 To 0) As Byte         'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    CR_LF(0 To 1) As Byte           'CR.LF
End Type

'データ・バッファ
Global HS_ZAIREC As HS_ZAIREC_Tag
Function HS_ZAI_Open1(Mode As Integer, FPass As String) As Integer
'********************************************************************
'*
'*       洗濯機事業部  在庫設定データ  ＯＰＥＮ
'*
'*      引数　:OPENモード（0:参照　1:更新）
'*
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.05.18
'*
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo HS_ZAI_Op_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    HS_ZAI_Open1 = False
                            'ホスト受信データフルパス取込み
    sts = GetIni("FILE", HS_ZAI_ID1, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        HS_ZAI_Open1 = True
        Exit Function
    End If
    FullPath = RTrim(c)
    FPass = FullPath

    HS_ZAI_No = FreeFile

    If Mode = ZERO Then
        Open FullPath For Input As #HS_ZAI_No
    Else
        Open FullPath For Binary As #HS_ZAI_No
    End If

    Exit Function

HS_ZAI_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
                ans = MsgBox("エラー [HS_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
    HS_ZAI_Open1 = True
    Exit Function
End Function
Function HS_ZAI_Get1() As Integer
'********************************************************************
'*
'*              在庫設定データ  ＧＥＴ
'*
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.05.18
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo HS_ZAI_Put_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    HS_ZAI_Get1 = False

    Get #HS_ZAI_No, , HS_ZAIREC

Exit Function

HS_ZAI_Put_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
            ans = MsgBox("ドライブまたはパスが見つかりません", vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [HS_ZAI Put : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    HS_ZAI_Get1 = True
    Exit Function
End Function
