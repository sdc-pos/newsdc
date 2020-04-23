Attribute VB_Name = "SYSHIN"
Option Explicit
'********************************************************************
'*                                                                  *
'*              旧システム品目データ（取込みワーク） ファイル定義            *
'*                                                                  *
'*          CREATE 1997.08.27  M.Yoshizawa                            *
'********************************************************************
'ファイルＩＤ
Global Const SYS_HIN_ID = "SYS_HIN"
'ファイル№
Global SYS_HIN_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SYS_HINREC_Tag
    No(0 To 3) As Byte               '
    JGYOBU(0 To 0) As Byte          '事業部区分
    NAIGAI(0 To 0) As Byte          '国内外
    HIN_GAI(0 To 12) As Byte        '品番（外部）
    HIN_NAME(0 To 24) As Byte       '品名
    ST_SET_DT(0 To 7) As Byte       '標準倉庫設定日付
    ST_SOKO(0 To 1) As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1) As Byte         '             列
    ST_REN(0 To 1) As Byte          '             連
    ST_DAN(0 To 1) As Byte          '             段
    BEF_SOKO(0 To 1) As Byte        '前回入庫倉庫 倉庫
    BEF_RETU(0 To 1) As Byte        '             列
    BEF_REN(0 To 1) As Byte         '             連
    BEF_DAN(0 To 1) As Byte         '             段
    LAST_NYU_DT(0 To 7) As Byte     '最終入庫日付
    LAST_SYU_DT(0 To 7) As Byte     '最終出庫日付
    HIN_NAI(0 To 12) As Byte        '品番（内部）
    BIKOU_SOKO(0 To 1) As Byte      '備考 ホスト倉庫
    BIKOU_TANA(0 To 7) As Byte      '備考 ホスト棚番
    SIZAI_CD(0 To 4) As Byte        '資材コード
    HOJYU_P(0 To 7) As Byte         '補充点
    AVE_SYUKA(0 To 7) As Byte       '月平均出荷数
    SAMPLE_QTY(0 To 0) As Byte       'サンプル数
    LAST_INP_DT(0 To 7) As Byte     '最終入荷日付
    FILLER(0 To 12) As Byte         'FILLER
End Type

'データ・バッファ
Global SYS_HINREC As SYS_HINREC_Tag
Function SYS_HIN_Open(Mode As Integer, FPass As String) As Integer
'********************************************************************
'*                                                                  *
'*      旧品目マスタデータ  ＯＰＥＮ         　                       *
'*                                                                  *
'*      引数　:OPENモード（0:参照　1:更新）                           *
'*                                                                  *
'*      戻り値:false 正常                                            *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.08.27  M.Yoshizawa                          *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo SYS_HIN_Op_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    SYS_HIN_Open = False
                            'ホスト受信データフルパス取込み
    sts = GetIni("FILE", SYS_HIN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SYS_HIN_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    FPass = FullPath

    SYS_HIN_No = FreeFile

    If Mode = 0 Then
        Open FullPath For Input As #SYS_HIN_No
    Else
        Open FullPath For Binary As #SYS_HIN_No
    End If

    Exit Function

SYS_HIN_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
                ans = MsgBox("エラー [WK_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
    SYS_HIN_Open = True
    Exit Function
End Function
Function SYS_HIN_Get() As Integer
'********************************************************************
'*                                                                  *
'*              旧品目マスタデータ（取込みワーク）  ＧＥＴ　           *
'*                                                                  *
'*      戻り値:false 正常                                            *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.08.26  M.Yoshizawa                          *
'********************************************************************
Dim ans As Integer
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

    On Error GoTo SYS_HIN_Get_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    SYS_HIN_Get = False

    Get #SYS_HIN_No, , SYS_HINREC

Exit Function

SYS_HIN_Get_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
            ans = MsgBox("ドライブまたはパスが見つかりません" & SYS_HIN_ID _
                  , vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [WK_ZAI Get : " & Str(Err.Number) & _
                  "] " & Err.Description, vbCritical)
    End Select
    SYS_HIN_Get = True
    Exit Function
End Function
