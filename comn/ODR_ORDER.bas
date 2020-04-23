Attribute VB_Name = "ODR_ORDER"
Option Explicit
'********************************************************************
'*                                                                  *
'*              親品番　注文Ｆ ファイル定義                           *
'*                                                                  *
'*          CREATE 2008.02.19                                       *
'
'           2012.04.13          PRT_FLG 追加
'
'********************************************************************
'ファイルＩＤ
Public Const ODR_ORDER_ID$ = "ODR_ORDER"

'ページサイズ
Private Const ODR_ORDER_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_ORDER_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_ORDER_REC_Tag
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    BUN_KB(0 To 2)              As Byte         '分納　有無区分
    REQ_KB(0 To 0)              As Byte         '展開区分
    ODR_QTY(0 To 4)             As Byte         '注文数
    CYUMON_DT(0 To 7)           As Byte         '部材センター注文納期（YYYYMMDD）
    KAITO_DT(0 To 7)            As Byte         '回答納期
    FIN_DT(0 To 7)              As Byte         '完了日付
    KUMI_OK_DT(0 To 7)          As Byte         '組立可能日付
    ODR_BMN(0 To 4)             As Byte         '発注部門
    DEN_NO(0 To 9)              As Byte         '伝票№
    UPD_TANTO(0 To 4)           As Byte         '更新　担当者
    INS_DT(0 To 7)              As Byte         '追加　日付
    INS_TM(0 To 5)              As Byte         '追加　時刻
    USE_YM_MOTO(0 To 5)         As Byte         '使用月（YYYYMM）プログラム起動時の内容
'    FILLER(0 To 21)             As Byte         'Filler
    UPD_DT(0 To 7)              As Byte         '更新　日付
    UPD_TM(0 To 5)              As Byte         '更新　時刻
    UPD_PG(0 To 6)              As Byte         '更新　プログラム
    PRT_FLG(0 To 0)             As Byte         '指図表印刷         F:済み、他:未印刷   2012.04.13

End Type
'データ・バッファ
Public ODR_ORDER_REC            As ODR_ORDER_REC_Tag



'キー定義

Type KEY0_ODR_ORDER                           'ＫＥＹ０
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
End Type

Type KEY1_ODR_ORDER                           'ＫＥＹ１
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
End Type

Type KEY2_ODR_ORDER                           'ＫＥＹ２
    ODR_QTY(0 To 4)             As Byte         '注文数
End Type

Type KEY3_ODR_ORDER                           'ＫＥＹ３
    KAITO_DT(0 To 7)            As Byte         '回答納期
End Type

Type KEY4_ODR_ORDER                           'ＫＥＹ４
    FIN_DT(0 To 7)              As Byte         '完了日付
End Type

Type KEY5_ODR_ORDER                           'ＫＥＹ５         '2009/03/12追加
    SHIMUKE(0 To 1)             As Byte         '仕向け先
    JGYOBU(0 To 0)              As Byte         '事業部
    NAIGAI(0 To 0)              As Byte         '国内外
    USE_YM(0 To 5)              As Byte         '使用月（YYYYMM)
    INS_DT(0 To 7)              As Byte         '追加　日付
    INS_TM(0 To 5)              As Byte         '追加　時刻
    HIN_GAI(0 To 19)            As Byte         '親品番
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
    INS_NO(0 To 3)              As Byte         '登録順
    BUN_NO(0 To 2)              As Byte         '分納回数
End Type

Type KEY6_ODR_ORDER                             'ＫＥＹ６         '20012/04/13追加
    ORDER_NO(0 To 9)            As Byte         '親品番　注文№
End Type


'キー・データ
Public K0_ODR_ORDER           As KEY0_ODR_ORDER
Public K1_ODR_ORDER           As KEY1_ODR_ORDER
Public K2_ODR_ORDER           As KEY2_ODR_ORDER

Public K3_ODR_ORDER           As KEY3_ODR_ORDER
Public K4_ODR_ORDER           As KEY4_ODR_ORDER
Public K5_ODR_ORDER           As KEY5_ODR_ORDER

Public K6_ODR_ORDER           As KEY6_ODR_ORDER


Type ODR_ORDER_FSpeck
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

    ks11                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private ODR_ORDER_Speck       As ODR_ORDER_FSpeck
Private Function ODR_ORDER_Create() As Integer
'********************************************************************
'*                                                                  *
'*              親＿注文Ｆ  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_ORDER_Create = True
                                            '親＿注文Ｆフルパス取込み
    sts = GetIni("FILE", ODR_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ODR_ORDER]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    ODR_ORDER_Speck.fs.recoleng = Len(ODR_ORDER_REC)      ' レコード長
    ODR_ORDER_Speck.fs.PageSize = ODR_ORDER_PG_SIZ        ' ページサイズ
    ODR_ORDER_Speck.fs.idexnumb = 7                       ' インデックス数
    ODR_ORDER_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_ORDER_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_ORDER_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_ORDER_Speck.ks0.keyleng = 41                      ' キー長
    ODR_ORDER_Speck.ks0.keyflag = BtKfChg + BtKfExt       ' キーフラグ
    ODR_ORDER_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks0.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー１ ▽
    ODR_ORDER_Speck.ks1.keypos = 1                        ' キーポジション
    ODR_ORDER_Speck.ks1.keyleng = 4                       ' キー長
    ODR_ORDER_Speck.ks1.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_ORDER_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks1.reserve = &H0                     ' 予約済み
    
    ODR_ORDER_Speck.ks2.keypos = 42                       ' キーポジション
    ODR_ORDER_Speck.ks2.keyleng = 6                       ' キー長
    ODR_ORDER_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_ORDER_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks2.reserve = &H0                     ' 予約済み
    
    ODR_ORDER_Speck.ks3.keypos = 5                        ' キーポジション
    ODR_ORDER_Speck.ks3.keyleng = 37                      ' キー長
    ODR_ORDER_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg      ' キーフラグ
    ODR_ORDER_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks3.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    ODR_ORDER_Speck.ks4.keypos = 52                       ' キーポジション
    ODR_ORDER_Speck.ks4.keyleng = 5                       ' キー長
    ODR_ORDER_Speck.ks4.keyflag = BtKfDup + BtKfChg + BtKfExt       ' キーフラグ
    ODR_ORDER_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks4.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー２ △
    
    
    '--------------------------------------------------- キー３ ▽
    
    ODR_ORDER_Speck.ks5.keypos = 65                         ' キーポジション
    ODR_ORDER_Speck.ks5.keyleng = 8                         ' キー長
                                                            ' キーフラグ
    ODR_ORDER_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_ORDER_Speck.ks5.keytype = Chr(BtKtString)           ' キータイプ
    ODR_ORDER_Speck.ks5.reserve = &H0                       ' 予約済み
    '--------------------------------------------------- キー３ △
    
    '--------------------------------------------------- キー４ ▽
    
    ODR_ORDER_Speck.ks6.keypos = 73                         ' キーポジション
    ODR_ORDER_Speck.ks6.keyleng = 8                         ' キー長
                                                            ' キーフラグ
    ODR_ORDER_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg
    ODR_ORDER_Speck.ks6.keytype = Chr(BtKtString)           ' キータイプ
    ODR_ORDER_Speck.ks6.reserve = &H0                       ' 予約済み
    '--------------------------------------------------- キー４ △
    
    '--------------------------------------------------- キー５ ▽
    ODR_ORDER_Speck.ks7.keypos = 1                        ' キーポジション
    ODR_ORDER_Speck.ks7.keyleng = 4                       ' キー長
    ODR_ORDER_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg       ' キーフラグ
    ODR_ORDER_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks7.reserve = &H0                     ' 予約済み
    
    ODR_ORDER_Speck.ks8.keypos = 42                       ' キーポジション
    ODR_ORDER_Speck.ks8.keyleng = 6                       ' キー長
    ODR_ORDER_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_ORDER_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks8.reserve = &H0                     ' 予約済み
    
    ODR_ORDER_Speck.ks9.keypos = 109                      ' キーポジション
    ODR_ORDER_Speck.ks9.keyleng = 14                      ' キー長
    ODR_ORDER_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg     ' キーフラグ
    ODR_ORDER_Speck.ks9.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks9.reserve = &H0                     ' 予約済み
    
    
    ODR_ORDER_Speck.ks10.keypos = 5                        ' キーポジション
    ODR_ORDER_Speck.ks10.keyleng = 37                      ' キー長
    ODR_ORDER_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg      ' キーフラグ
    ODR_ORDER_Speck.ks10.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ORDER_Speck.ks10.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー５ △
    '--------------------------------------------------- キー６ ▽
    ODR_ORDER_Speck.ks11.keypos = 25                        ' キーポジション
    ODR_ORDER_Speck.ks11.keyleng = 10                       ' キー長
                                                            ' キーフラグ
    ODR_ORDER_Speck.ks11.keyflag = BtKfChg + BtKfDup + BtKfExt
    ODR_ORDER_Speck.ks11.keytype = Chr(BtKtString)          ' キータイプ
    ODR_ORDER_Speck.ks11.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー６ △
    
    sts = BTRV(BtOpCreate, ODR_ORDER_POS, ODR_ORDER_Speck, Len(ODR_ORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "親＿注文Ｆ")
        Exit Function
    End If
    
    ODR_ORDER_Create = False

End Function

Public Function ODR_ORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              親＿注文Ｆ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ODR_ORDER_Open = True
                                            '親＿注文Ｆフルパス取込み
    sts = GetIni("FILE", ODR_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ODR_ORDER]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ODR_ORDER_Create()      '親＿注文Ｆ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "親 注文Ｆ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "親 注文Ｆ")
                Exit Function
        End Select
    Loop
    
    ODR_ORDER_Open = False
    
End Function
Public Function ODR_ORDER_GET(SM As String, JB As String, NG As String, HG As String, _
                               i_NO As String, OD As String, BN As String, Locked As Integer) As Integer
'           引数
'   JB      事業部
'   NG      内外
'   HG      品番
'   OD      注文№
'   I_No    Key　№
'   BN      分納回数
'   Locked  ＧｅｔＬｏｃｋ


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_ORDER_GET = True
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, SM)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, JB)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, NG)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, HG)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, i_NO)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, OD)
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, BN)
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'Beep
                'MsgBox "指定された工程がありません。"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Function
        End Select
    Loop
    
    ODR_ORDER_GET = False

End Function


Public Sub ODR_ORDER_CLR()
    
    Call UniCode_Conv(ODR_ORDER_REC.SHIMUKE, "")
    Call UniCode_Conv(ODR_ORDER_REC.JGYOBU, "")
    Call UniCode_Conv(ODR_ORDER_REC.NAIGAI, "")
    Call UniCode_Conv(ODR_ORDER_REC.USE_YM, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_NO, String(UBound(ODR_ORDER_REC.INS_NO) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, "")
    Call UniCode_Conv(ODR_ORDER_REC.BUN_KB, String(UBound(ODR_ORDER_REC.BUN_KB) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.REQ_KB, String(UBound(ODR_ORDER_REC.REQ_KB) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, String(UBound(ODR_ORDER_REC.ODR_QTY) + 1, "0"))
    Call UniCode_Conv(ODR_ORDER_REC.CYUMON_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.KAITO_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.FIN_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.KUMI_OK_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.ODR_BMN, "")
    Call UniCode_Conv(ODR_ORDER_REC.DEN_NO, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_TANTO, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.INS_TM, "")
    Call UniCode_Conv(ODR_ORDER_REC.USE_YM_MOTO, "")
    'Call UniCode_Conv(ODR_ORDER_REC.FILLER, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_DT, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_TM, "")
    Call UniCode_Conv(ODR_ORDER_REC.UPD_PG, "")
    Call UniCode_Conv(ODR_ORDER_REC.PRT_FLG, "")

End Sub

