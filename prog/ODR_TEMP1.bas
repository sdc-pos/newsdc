Attribute VB_Name = "ODR_TEMP1"
Option Explicit
'********************************************************************
'*                                                                  *
'*              中間　所要量Ｆ（WORK) ファイル定義              　　　*
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_TEMP1_ID$ = "ODR_TEMP1"

'ページサイズ
Private Const ODR_TEMP1_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_TP1_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_TP1_R_Tag

    KAITO_DT(0 To 7)            As Byte         '親注文の回答納期
    CYUMON_DT(0 To 7)           As Byte         '部材センター注文納期（YYYYMMDD）
    USE_YM(0 To 5)              As Byte         '使用月
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    KO_SYUBETSU(0 To 1)         As Byte         '子　種別
    KO_QTY(0 To 5)              As Byte         '子　員数(999V99)
    OK_DT(0 To 7)               As Byte         '出庫可能日 YYYYMMDD
    KAN_KB(0 To 0)              As Byte         '親注文の完成区分
    ALL_QTY(0 To 8)             As Byte         '展開数     9(5)v9(2)
    USE_QTY(0 To 8)             As Byte         '使用数     9(5)v9(2)
    NED_QTY(0 To 8)             As Byte         '必要数     9(5)v9(2)
    REQ_QTY(0 To 8)             As Byte         '所要数     9(5)v9(2)
    FUSOKU_QTY(0 To 8)          As Byte         '不足数     9(5)v9(2)
    UPDT_DT(0 To 5)             As Byte         '更新日     YYMMDD
    UPDT_TM(0 To 3)             As Byte         '更新時刻   hhmm
    FILLER(0 To 17)             As Byte         'Filler

End Type
'データ・バッファ
Public ODR_TP1_R            As ODR_TP1_R_Tag



'キー定義

Type KEY0_ODR_TEMP1                           'ＫＥＹ０

    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    
End Type

Type KEY1_ODR_TEMP1                           'ＫＥＹ１

    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    
    OK_DT(0 To 7)               As Byte         '出庫可能日 YYYYMMDD

End Type

Type KEY2_ODR_TEMP1                           'ＫＥＹ２

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数

End Type

Type KEY3_ODR_TEMP1                           'ＫＥＹ３

    KAN_KB(0 To 0)              As Byte         '親注文の完成区分
    
    KAITO_DT(0 To 7)            As Byte         '親注文の回答納期
    CYUMON_DT(0 To 7)           As Byte         '部材センター注文納期（YYYYMMDD）
    USE_YM(0 To 5)              As Byte         '使用月
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番

End Type

Type KEY4_ODR_TEMP1                           'ＫＥＹ４（2010/05/07追加）
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    
    KAN_KB(0 To 0)              As Byte         '親注文の完成区分
    
    KAITO_DT(0 To 7)            As Byte         '親注文の回答納期
    CYUMON_DT(0 To 7)           As Byte         '部材センター注文納期（YYYYMMDD）
    USE_YM(0 To 5)              As Byte         '使用月
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数

End Type

'キー・データ
Public K0_ODR_TEMP1           As KEY0_ODR_TEMP1
Public K1_ODR_TEMP1           As KEY1_ODR_TEMP1
Public K2_ODR_TEMP1           As KEY2_ODR_TEMP1
Public K3_ODR_TEMP1           As KEY3_ODR_TEMP1
Public K4_ODR_TEMP1           As KEY4_ODR_TEMP1

Type ODR_TEMP1_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks7                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks8                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks9                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private ODR_TEMP1_Speck       As ODR_TEMP1_FSpeck
Private Function ODR_TEMP1_Create() As Integer
'********************************************************************
'*                                                                  *
'*              中間所要量Ｆ  ＣＲＥＡＴＥ                            *
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

    ODR_TEMP1_Create = True
                                            '中間所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]読み込みエラー")
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


    ODR_TEMP1_Speck.fs.recoleng = Len(ODR_TP1_R)      ' レコード長
    ODR_TEMP1_Speck.fs.PageSize = ODR_TEMP1_PG_SIZ          ' ページサイズ
    ODR_TEMP1_Speck.fs.idexnumb = 5                       ' インデックス数
    ODR_TEMP1_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_TEMP1_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_TEMP1_Speck.ks0.keypos = 23                       ' キーポジション
    ODR_TEMP1_Speck.ks0.keyleng = 63                      ' キー長
    ODR_TEMP1_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP1_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    ODR_TEMP1_Speck.ks1.keypos = 23                        ' キーポジション
    ODR_TEMP1_Speck.ks1.keyleng = 41                      ' キー長
    ODR_TEMP1_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt    ' キーフラグ
    ODR_TEMP1_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks1.reserve = &H0                     ' 予約済み
    
    ODR_TEMP1_Speck.ks2.keypos = 89                       ' キーポジション
    ODR_TEMP1_Speck.ks2.keyleng = 8                       ' キー長
    ODR_TEMP1_Speck.ks2.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP1_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks2.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    ODR_TEMP1_Speck.ks3.keypos = 64                       ' キーポジション
    ODR_TEMP1_Speck.ks3.keyleng = 22                      ' キー長
    ODR_TEMP1_Speck.ks3.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' キーフラグ
    ODR_TEMP1_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks3.reserve = &H0                     ' 予約済み
    
    ODR_TEMP1_Speck.ks4.keypos = 23                       ' キーポジション
    ODR_TEMP1_Speck.ks4.keyleng = 41                      ' キー長
    ODR_TEMP1_Speck.ks4.keyflag = BtKfChg + BtKfDup + BtKfExt               ' キーフラグ
    ODR_TEMP1_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks4.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー２ △
    
    '--------------------------------------------------- キー３ ▽
    ODR_TEMP1_Speck.ks5.keypos = 102                     ' キーポジション
    ODR_TEMP1_Speck.ks5.keyleng = 1                      ' キー長
    ODR_TEMP1_Speck.ks5.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' キーフラグ
    ODR_TEMP1_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks5.reserve = &H0                     ' 予約済み
    
    ODR_TEMP1_Speck.ks6.keypos = 1                        ' キーポジション
    ODR_TEMP1_Speck.ks6.keyleng = 85                       ' キー長
    ODR_TEMP1_Speck.ks6.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP1_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks6.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー３ △
    
    
    '--------------------------------------------------- キー４ ▽      '2010/05/07追加
    ODR_TEMP1_Speck.ks7.keypos = 64                       ' キーポジション
    ODR_TEMP1_Speck.ks7.keyleng = 22                      ' キー長
    ODR_TEMP1_Speck.ks7.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' キーフラグ
    ODR_TEMP1_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks7.reserve = &H0                     ' 予約済み
    
    ODR_TEMP1_Speck.ks8.keypos = 102                     ' キーポジション
    ODR_TEMP1_Speck.ks8.keyleng = 1                      ' キー長
    ODR_TEMP1_Speck.ks8.keyflag = BtKfChg + BtKfDup + BtKfSeg + BtKfExt       ' キーフラグ
    ODR_TEMP1_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks8.reserve = &H0                     ' 予約済み
    
    ODR_TEMP1_Speck.ks9.keypos = 1                        ' キーポジション
    ODR_TEMP1_Speck.ks9.keyleng = 63                       ' キー長
    ODR_TEMP1_Speck.ks9.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP1_Speck.ks9.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP1_Speck.ks9.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー４ △
    

    sts = BTRV(BtOpCreate, ODR_TP1_POS, ODR_TEMP1_Speck, Len(ODR_TEMP1_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP1")
        Exit Function
    End If
    
    ODR_TEMP1_Create = False

End Function

Public Function ODR_TEMP1_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              所要量Ｆ  ＯＰＥＮ
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

    ODR_TEMP1_Open = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]読み込みエラー")
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
        sts = BTRV(BtOpOpen, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP1_Create()      '所要量Ｆ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP1")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP1")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP1_Open = False
    
End Function

Public Function ODR_TEMP1_KILL() As Integer
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

    ODR_TEMP1_KILL = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP1_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP1]読み込みエラー")
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
    
    ODR_TEMP1_KILL = False
    
End Function

Public Function ODR_TEMP1_GET(SM As String, JB As String, NG As String, _
                    YM As String, HG As String, OD As String, BN As String, Locked As Integer) As Integer
'           引数
'   SM      仕向先
'   JB      事業部
'   NG      内外
'   YM      使用月
'   HG      親品番
'   OD      注文№
'   BN      分納回数
'   Locked  ＧｅｔＬｏｃｋ


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_TEMP1_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP1.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_TEMP1.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_TEMP1.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_TEMP1.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_TEMP1.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_TEMP1.BUN_NO, BN)
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP1_GET = False

End Function

Public Sub ODR_TEMP1_CLR()
    
    Call UniCode_Conv(ODR_TP1_R.KAITO_DT, "")
    Call UniCode_Conv(ODR_TP1_R.CYUMON_DT, "")
    Call UniCode_Conv(ODR_TP1_R.USE_YM, "")
    Call UniCode_Conv(ODR_TP1_R.SHIMUKE, "")
    Call UniCode_Conv(ODR_TP1_R.JGYOBU, "")
    Call UniCode_Conv(ODR_TP1_R.NAIGAI, "")
    Call UniCode_Conv(ODR_TP1_R.HIN_GAI, "")
    Call UniCode_Conv(ODR_TP1_R.ORDER_NO, "")
    Call UniCode_Conv(ODR_TP1_R.INS_NO, "")
    Call UniCode_Conv(ODR_TP1_R.BUN_NO, "")
    
    Call UniCode_Conv(ODR_TP1_R.KO_JGYOBU, "")
    Call UniCode_Conv(ODR_TP1_R.KO_NAIGAI, "")
    Call UniCode_Conv(ODR_TP1_R.KO_HIN_GAI, "")
    Call UniCode_Conv(ODR_TP1_R.KO_SYUBETSU, "")
    Call UniCode_Conv(ODR_TP1_R.KO_QTY, String(UBound(ODR_TP1_R.KO_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.OK_DT, "")
    Call UniCode_Conv(ODR_TP1_R.KAN_KB, "1")
    Call UniCode_Conv(ODR_TP1_R.REQ_QTY, String(UBound(ODR_TP1_R.REQ_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.USE_QTY, String(UBound(ODR_TP1_R.USE_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.NED_QTY, String(UBound(ODR_TP1_R.NED_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.FUSOKU_QTY, String(UBound(ODR_TP1_R.FUSOKU_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_TP1_R.UPDT_DT, "")
    
    Call UniCode_Conv(ODR_TP1_R.UPDT_TM, "")
    Call UniCode_Conv(ODR_TP1_R.FILLER, "")
    
End Sub

