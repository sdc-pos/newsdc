Attribute VB_Name = "ODR_TEMP2"
Option Explicit
'********************************************************************
'*                                                                  *
'*              中間　所要量Ｆ（WORK) ファイル定義              　　　*
'*                                                                  *
'*          CREATE 2008.03.06                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_TEMP2_ID$ = "ODR_TEMP2"

'ページサイズ
Private Const ODR_TEMP2_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_TP2_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_TP2_R_Tag


    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    IO_KB(0 To 0)               As Byte         'io区分
    USE_YM(0 To 5)              As Byte         '使用月
    ANS_NOUKI_DT(0 To 7)        As Byte         '対象日付   YYYYMMDD    (回答納期）
    ORDER_NO(0 To 4)            As Byte         '注文№
    ZAI_QTY(0 To 8)             As Byte         '在庫数／発注数         9(5)v9(2)
    MOTO_QTY(0 To 8)            As Byte         '元々の在庫数／発注数 9(5)v9(2)
    UPDT_DT(0 To 5)             As Byte         '更新日     YYMMDD
    UPDT_TM(0 To 3)             As Byte         '更新時刻   hhmm
    FILLER(0 To 7)              As Byte         'Filler

End Type
'データ・バッファ
Public ODR_TP2_R            As ODR_TP2_R_Tag



'キー定義

Type KEY0_ODR_TEMP2                           'ＫＥＹ０

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    IO_KB(0 To 0)               As Byte         'io区分
    USE_YM(0 To 5)              As Byte         '使用月
    ANS_NOUKI_DT(0 To 7)        As Byte         '対象日付   YYYYMMDD    (回答納期）
    ORDER_NO(0 To 4)            As Byte         '注文№
    
End Type

Type KEY1_ODR_TEMP2                           'ＫＥＹ１

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    USE_YM(0 To 5)              As Byte         '使用月
    IO_KB(0 To 0)               As Byte         'io区分
    
End Type
'キー・データ
Public K0_ODR_TEMP2           As KEY0_ODR_TEMP2
Public K1_ODR_TEMP2           As KEY1_ODR_TEMP2

Type ODR_TEMP2_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private ODR_TEMP2_Speck       As ODR_TEMP2_FSpeck
Private Function ODR_TEMP2_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ODR_TEMP2  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_TEMP2_Create = True
                                            'ODR_TEMP2 フルパス取込み
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_STR



    ODR_TEMP2_Speck.fs.recoleng = Len(ODR_TP2_R)      ' レコード長
    ODR_TEMP2_Speck.fs.PageSize = ODR_TEMP2_PG_SIZ          ' ページサイズ
    ODR_TEMP2_Speck.fs.idexnumb = 2                       ' インデックス数
    ODR_TEMP2_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_TEMP2_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_TEMP2_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_TEMP2_Speck.ks0.keyleng = 42                      ' キー長
    ODR_TEMP2_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP2_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP2_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    '--------------------------------------------------- キー１ ▽
    ODR_TEMP2_Speck.ks1.keypos = 1                        ' キーポジション
    ODR_TEMP2_Speck.ks1.keyleng = 22                      ' キー長
    ODR_TEMP2_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt     ' キーフラグ
    ODR_TEMP2_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP2_Speck.ks1.reserve = &H0                     ' 予約済み
    
    ODR_TEMP2_Speck.ks2.keypos = 24                       ' キーポジション
    ODR_TEMP2_Speck.ks2.keyleng = 6                       ' キー長
    ODR_TEMP2_Speck.ks2.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt     ' キーフラグ
    ODR_TEMP2_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP2_Speck.ks2.reserve = &H0                     ' 予約済み
    
    ODR_TEMP2_Speck.ks3.keypos = 23                       ' キーポジション
    ODR_TEMP2_Speck.ks3.keyleng = 1                       ' キー長
    ODR_TEMP2_Speck.ks3.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP2_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP2_Speck.ks3.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー１ △
    
    

    sts = BTRV(BtOpCreate, ODR_TP2_POS, ODR_TEMP2_Speck, Len(ODR_TEMP2_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP2")
        Exit Function
    End If
    
    ODR_TEMP2_Create = False

End Function

Public Function ODR_TEMP2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_TEMP2  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim yn          As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_TEMP2_Open = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)


    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_STR



    Do
        sts = BTRV(BtOpOpen, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("他で使用中です！<ODR_TEMP2>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP2_Create()      'ODR_TEMP2 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP2")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP2_Open = False
    
End Function

Public Function ODR_TEMP2_KILL() As Integer
'********************************************************************
'*
'*              所要量Ｆ  削除＆再作成（Ｏｐｅｎ）
'*
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_TEMP2_KILL = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP2_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP2]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If

    W_STR = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)

    FullPath = W_STR
    
    Kill FullPath
    
    ODR_TEMP2_KILL = False
    
End Function

Public Function ODR_TEMP2_GET(JB As String, NG As String, HG As String, _
                    Kb As String, DT As String, OD As String, Locked As Integer) As Integer
'           引数

'   JB      事業部
'   NG      内外
'   HG      子品番
'   KB      io区分
'   DT      納期
'   OD      注文№

'   Locked  ＧｅｔＬｏｃｋ
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_TEMP2_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, JB)       '子　事業部
    Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, NG)       '子　国内外
    Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, HG)      '子品番
    Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, Kb)           'io区分
    Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, DT)       '注文日
    Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, OD)        '注文№
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'MsgBox "指定された工程がありません。"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<ODR_TEMP2>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP2")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP2_GET = False

End Function
    

Public Sub ODR_TEMP2_CLR()
    '子　事業部
    Call UniCode_Conv(ODR_TP2_R.KO_JGYOBU, "")
    '子　国内外
    Call UniCode_Conv(ODR_TP2_R.KO_NAIGAI, "")
    '子品番
    Call UniCode_Conv(ODR_TP2_R.KO_HIN_GAI, "")
    'io区分
    Call UniCode_Conv(ODR_TP2_R.IO_KB, "")
    '使用月
    Call UniCode_Conv(ODR_TP2_R.USE_YM, "")
    '納期
    Call UniCode_Conv(ODR_TP2_R.ANS_NOUKI_DT, "")
    '注文№
    Call UniCode_Conv(ODR_TP2_R.ORDER_NO, "")
    '在庫数     9(5)v9(2)
    Call UniCode_Conv(ODR_TP2_R.ZAI_QTY, String(UBound(ODR_TP2_R.ZAI_QTY) + 1, "0"))
    '在庫数     9(5)v9(2)
    Call UniCode_Conv(ODR_TP2_R.MOTO_QTY, String(UBound(ODR_TP2_R.MOTO_QTY) + 1, "0"))
    '更新日     yymmdd
    Call UniCode_Conv(ODR_TP2_R.UPDT_DT, "")
    '更新時刻   hhmm
    Call UniCode_Conv(ODR_TP2_R.UPDT_TM, "")
    
    Call UniCode_Conv(ODR_TP2_R.FILLER, "")
End Sub

