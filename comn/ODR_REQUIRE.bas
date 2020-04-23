Attribute VB_Name = "ODR_REQUIRE"
Option Explicit
'********************************************************************
'*                                                                  *
'*              所要量Ｆ ファイル定義                           　　　*
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_REQUIRE_ID$ = "ODR_REQUIRE"

'ページサイズ
Private Const ODR_REQUIRE_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_REQ_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_REQ_R_Tag

    SHIMUKE(0 To 1)        As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    KO_SYUBETSU(0 To 1)         As Byte         '子　種別
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    CYUMON_DT(0 To 7)           As Byte         '部材センター注文納期（YYYYMMDD）
    REQ_QTY(0 To 7)             As Byte         '展開数     9(5)v9(2)
    ODR_QTY(0 To 7)             As Byte         '所要数     9(5)v9(2)
    FUSOKU_QTY(0 To 7)          As Byte         '不足数     9(5)v9(2)
    UPD_TANTO(0 To 4)           As Byte         '更新　担当者
    UPD_DT(0 To 7)              As Byte         '更新　日
    UPD_TM(0 To 5)              As Byte         '更新　時刻
    OK_DT(0 To 7)               As Byte         '
    FILLER(0 To 19)             As Byte         'Filler

End Type
'データ・バッファ
Public ODR_REQ_R            As ODR_REQ_R_Tag



'キー定義

Type KEY0_ODR_REQUIRE                           'ＫＥＹ０

    SHIMUKE(0 To 1)        As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
     
End Type

Type KEY1_ODR_REQUIRE                           'ＫＥＹ１

    SHIMUKE(0 To 1)        As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数

End Type

Type KEY2_ODR_REQUIRE                           'ＫＥＹ２

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    SHIMUKE(0 To 1)        As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数

End Type

Type KEY3_ODR_REQUIRE                           'ＫＥＹ３

    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    SHIMUKE(0 To 1)        As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数

End Type


'キー・データ
Public K0_ODR_REQ           As KEY0_ODR_REQUIRE
Public K1_ODR_REQ           As KEY1_ODR_REQUIRE
Public K2_ODR_REQ           As KEY2_ODR_REQUIRE
Public K3_ODR_REQ           As KEY3_ODR_REQUIRE

Type ODR_REQUIRE_FSpeck
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
    ks10                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private ODR_REQUIRE_Speck       As ODR_REQUIRE_FSpeck
Private Function ODR_REQUIRE_Create() As Integer
'********************************************************************
'*                                                                  *
'*              所要量Ｆ  ＣＲＥＡＴＥ                               *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_REQUIRE_Create = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_REQUIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_REQUIRE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    ODR_REQUIRE_Speck.fs.recoleng = Len(ODR_REQ_R)      ' レコード長
    ODR_REQUIRE_Speck.fs.PageSize = ODR_REQUIRE_PG_SIZ          ' ページサイズ
    ODR_REQUIRE_Speck.fs.idexnumb = 4                       ' インデックス数
    ODR_REQUIRE_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_REQUIRE_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_REQUIRE_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_REQUIRE_Speck.ks0.keyleng = 61                      ' キー長
    ODR_REQUIRE_Speck.ks0.keyflag = BtKfChg + BtKfExt       ' キーフラグ
    ODR_REQUIRE_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    '--------------------------------------------------- キー１ ▽
    ODR_REQUIRE_Speck.ks1.keypos = 1                        ' キーポジション
    ODR_REQUIRE_Speck.ks1.keyleng = 4                       ' キー長
    ODR_REQUIRE_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_REQUIRE_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks1.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks2.keypos = 42                       ' キーポジション
    ODR_REQUIRE_Speck.ks2.keyleng = 20                      ' キー長
    ODR_REQUIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_REQUIRE_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks2.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks3.keypos = 15                        ' キーポジション
    ODR_REQUIRE_Speck.ks3.keyleng = 37                      ' キー長
    ODR_REQUIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg      ' キーフラグ
    ODR_REQUIRE_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks3.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    ODR_REQUIRE_Speck.ks4.keypos = 64                        ' キーポジション
    ODR_REQUIRE_Speck.ks4.keyleng = 2                       ' キー長
    ODR_REQUIRE_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_REQUIRE_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks4.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks5.keypos = 42                       ' キーポジション
    ODR_REQUIRE_Speck.ks5.keyleng = 20                      ' キー長
    ODR_REQUIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_REQUIRE_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks5.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks6.keypos = 1                        ' キーポジション
    ODR_REQUIRE_Speck.ks6.keyleng = 41                      ' キー長
    ODR_REQUIRE_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg      ' キーフラグ
    ODR_REQUIRE_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks6.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー２ △
    '--------------------------------------------------- キー３ ▽
    ODR_REQUIRE_Speck.ks7.keypos = 66                        ' キーポジション
    ODR_REQUIRE_Speck.ks7.keyleng = 6                       ' キー長
    ODR_REQUIRE_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_REQUIRE_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks7.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks8.keypos = 64                        ' キーポジション
    ODR_REQUIRE_Speck.ks8.keyleng = 2                       ' キー長
    ODR_REQUIRE_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_REQUIRE_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks8.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks9.keypos = 42                       ' キーポジション
    ODR_REQUIRE_Speck.ks9.keyleng = 20                      ' キー長
    ODR_REQUIRE_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_REQUIRE_Speck.ks9.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks9.reserve = &H0                     ' 予約済み
    
    ODR_REQUIRE_Speck.ks10.keypos = 1                        ' キーポジション
    ODR_REQUIRE_Speck.ks10.keyleng = 41                      ' キー長
    ODR_REQUIRE_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg      ' キーフラグ
    ODR_REQUIRE_Speck.ks10.keytype = Chr(BtKtString)         ' キータイプ
    ODR_REQUIRE_Speck.ks10.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー３ △
    
    sts = BTRV(BtOpCreate, ODR_REQ_POS, ODR_REQUIRE_Speck, Len(ODR_REQUIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "所要量Ｆ")
        Exit Function
    End If
    
    ODR_REQUIRE_Create = False

End Function

Public Function ODR_REQUIRE_Open(Mode As Integer) As Integer
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

    ODR_REQUIRE_Open = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_REQUIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_REQUIRE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)


    Do
        sts = BTRV(BtOpOpen, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_REQUIRE_Create()      '所要量Ｆ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "所要量Ｆ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "所要量Ｆ")
                Exit Function
        End Select
    Loop
    
    ODR_REQUIRE_Open = False
    
End Function

Public Function ODR_REQUIRE_GET(SM As String, JB As String, NG As String, _
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

    ODR_REQUIRE_GET = True
    
    Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_REQ.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_REQ.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_REQ.BUN_NO, BN)
    
'2019.01.08    com = BtOpGetEqual + Locked
    com = BtOpGetEqual
    Do
        sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_REQUIRE")
                Exit Function
        End Select
    Loop
    
    ODR_REQUIRE_GET = False

End Function


Public Sub ODR_REQUIRE_CLR()
    
    Call UniCode_Conv(ODR_REQ_R.SHIMUKE, "")
    Call UniCode_Conv(ODR_REQ_R.JGYOBU, "")
    Call UniCode_Conv(ODR_REQ_R.NAIGAI, "")
    Call UniCode_Conv(ODR_REQ_R.USE_YM, "")
    Call UniCode_Conv(ODR_REQ_R.ORDER_NO, "")
    Call UniCode_Conv(ODR_REQ_R.INS_NO, "")
    Call UniCode_Conv(ODR_REQ_R.BUN_NO, "")
    Call UniCode_Conv(ODR_REQ_R.HIN_GAI, "")
    Call UniCode_Conv(ODR_REQ_R.REQ_QTY, String(UBound(ODR_REQ_R.REQ_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.ODR_QTY, String(UBound(ODR_REQ_R.ODR_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.FUSOKU_QTY, String(UBound(ODR_REQ_R.FUSOKU_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_REQ_R.KO_HIN_GAI, "")
    Call UniCode_Conv(ODR_REQ_R.KO_SYUBETSU, "")
    Call UniCode_Conv(ODR_REQ_R.KO_JGYOBU, "")
    Call UniCode_Conv(ODR_REQ_R.KO_NAIGAI, "")
    Call UniCode_Conv(ODR_REQ_R.CYUMON_DT, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_TANTO, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_DT, "")
    Call UniCode_Conv(ODR_REQ_R.UPD_TM, "")
    Call UniCode_Conv(ODR_REQ_R.OK_DT, "")
    Call UniCode_Conv(ODR_REQ_R.FILLER, "")
End Sub

