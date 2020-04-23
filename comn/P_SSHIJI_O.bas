Attribute VB_Name = "P_SSHIJI_O"
Option Explicit

'********************************************************************
'*
'*              商品化指図データ（親）  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SSHIJI_O_ID$ = "P_SSHIJI_O"

'ページサイズ
Private Const P_SSHIJI_O_PG_SIZ% = 2048

'ポジション・ブロック
Public P_SSHIJI_O_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

Private Type GENKA_TBL_Tag          '原価情報のﾃｰﾌﾞﾙ
    NIN(0 To 2)             As Byte         '人数
    TIMES(0 To 5)           As Byte         '時間
End Type




'レコード定義
Public Type P_SSHIJI_O_REC_Tag
    
    xSHIJI_NO(0 To 4)       As Byte         '指図票№   未使用とする 2007.11.28
    HAKKO_DT(0 To 7)        As Byte         '発行日
    Print_datetime(0 To 13) As Byte         '発行日時
    TANTO_CODE(0 To 4)      As Byte         '担当者ｺｰﾄﾞ
    SHONIN_CODE(0 To 4)     As Byte         '承認者ｺｰﾄﾞ
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    SHIJI_QTY(0 To 10)      As Byte         '指示数(9(8)V99)
    UKEHARAI_CODE(0 To 4)   As Byte         '手配先ｺｰﾄﾞ
    S_CLASS_CODE(0 To 19)   As Byte         '商品化ｸﾗｽ
    F_CLASS_CODE(0 To 19)   As Byte         '付加ｸﾗｽ
    N_CLASS_CODE(0 To 19)   As Byte         '内職ｸﾗｽ
    S_TANTO(0 To 1)         As Byte         '収単／担当者コード
    SAMPLE_F(0 To 0)        As Byte         '見本作成
    SHIJI_F(0 To 0)         As Byte         '指示形態 0:通常　1:ｽﾎﾟｯﾄ　2：欠品解除 3:再梱包(2007.11.09)
    TORI_KBN(0 To 0)        As Byte         '取引先コード
    
    PRI_SHIJI(0 To 0)       As Byte         '出力対象 指図票
    PRI_PARTS(0 To 0)       As Byte         '出力対象 ﾊﾟｰﾂﾗﾍﾞﾙ
    PRI_GAISOU(0 To 0)      As Byte         '出力対象 外装ﾗﾍﾞﾙ
    PRI_KISHU(0 To 0)       As Byte         '出力対象 機種ﾗﾍﾞﾙ
    
    BIKOU(0 To 119)         As Byte         '備考
    
    
    KAN_F(0 To 0)           As Byte         '完了F
    KAN_DT(0 To 7)          As Byte         '完了日
    BUNNOU_CNT(0 To 1)      As Byte         '分納回数
    UKEIRE_QTY(0 To 10)     As Byte         '受入数（合計）
                                            '原価項目
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '自責要因名
    JISEKI_NIN(0 To 2)      As Byte         '自責  人
    JISEKI_TIMES(0 To 5)    As Byte         '自責  分
    TASEKI_NAME(0 To 19)    As Byte         '他責要因名
    TASEKI_NIN(0 To 2)      As Byte         '他責  人
    TASEKI_TIMES(0 To 5)    As Byte         '他責  分
    
    
    CANCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
    
    ORDER_DT(0 To 7)        As Byte         '受注日(注文№) 2007.02.20
    
    
    SHIJI_No(0 To 7)        As Byte         '指図票№   未使用とする 2007.11.28
    
    
    HIN_CHECK_TANTO(0 To 4) As Byte         '品番ﾁｪｯｸ担当者ｺｰﾄﾞ 2010.09.03
    HIN_CHECK_DATETIME(0 To 13) _
                            As Byte         '品番ﾁｪｯｸ日時 2010.09.03
    HIN_CHECK_LABEL_CNT(0 To 2) _
                            As Byte         '品番ﾁｪｯｸﾗﾍﾞﾙ件数   2010.09.03
    HIN_CHECK_GENPIN_CNT(0 To 2) _
                            As Byte         '品番ﾁｪｯｸ現品票件数   2010.09.03
            
    ORDER_DT_SEQ(0 To 1)    As Byte         '受注日(注文№)枝番 2012.03.27
            
    COMPO_END_F(0 To 0)     As Byte         '構成ﾁｪｯｸ完了F(大阪PC) 9:完了 2012.04.13
    
'    FILLER(0 To 2)          As Byte         'Filler 2011.04.13         --> 2015.11.07
    
    HIN_CHECK_GAISOU_CNT(0 To 2) _
                            As Byte         '品番ﾁｪｯｸ外装品番件数   2015.11.07
    
    
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SSHIJI_O_REC       As P_SSHIJI_O_REC_Tag

'キー定義

Type KEY0_P_SSHIJI_O                        'ＫＥＹ０
'    SHIJI_NO(0 To 4)        As Byte         '指図票№
    SHIJI_No(0 To 7)        As Byte         '指図票№   2007.11.28
End Type

Type KEY1_P_SSHIJI_O                        'ＫＥＹ１
    KAN_F(0 To 0)           As Byte         '完了F
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    KAN_DT(0 To 7)          As Byte         '完了日
'    SHIJI_NO(0 To 4)       As Byte         '指図票№
    SHIJI_No(0 To 7)        As Byte         '指図票№   2007.11.28
End Type
    
Type KEY2_P_SSHIJI_O                        'ＫＥＹ２
    ORDER_DT(0 To 7)        As Byte         '受注日 2007.02.20
End Type
    
Type KEY3_P_SSHIJI_O                        'ＫＥＹ３   2007.11.14
    HAKKO_DT(0 To 7)        As Byte         '発行日
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    UKEHARAI_CODE(0 To 4)   As Byte         '手配先ｺｰﾄﾞ
End Type
    
Type KEY4_P_SSHIJI_O                        'ＫＥＹ４   2011.11.11
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    Print_datetime(0 To 13) As Byte         '発行日時
End Type
    
    
Type KEY5_P_SSHIJI_O                        'ＫＥＹ５   2012.03.08
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    KAN_F(0 To 0)           As Byte         '完了F
End Type
    
Type KEY6_P_SSHIJI_O                        'ＫＥＹ６   2012.03.27
    ORDER_DT(0 To 7)        As Byte         '受注日(注文№)     2007.02.20
    ORDER_DT_SEQ(0 To 1)    As Byte         '受注日(注文№)枝番 2012.03.27
End Type
    
    
    
    
    
'キー・データ
Public K0_P_SSHIJI_O        As KEY0_P_SSHIJI_O
Public K1_P_SSHIJI_O        As KEY1_P_SSHIJI_O
Public K2_P_SSHIJI_O        As KEY2_P_SSHIJI_O
Public K3_P_SSHIJI_O        As KEY3_P_SSHIJI_O      '2007.11.14

Public K4_P_SSHIJI_O        As KEY4_P_SSHIJI_O      '2011.11.11

Public K5_P_SSHIJI_O        As KEY5_P_SSHIJI_O      '2012.03.08

Public K6_P_SSHIJI_O        As KEY6_P_SSHIJI_O      '2012.03.27



Type P_SSHIJI_O_FSpeck
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

    '2011.11.11
    ks12                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks13                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks14                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks15                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks16                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    '2011.11.11

    '2012.03.08
    ks17                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks18                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks19                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks20                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks21                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    '2012.03.08


    '2012.03.27
    ks22                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks23                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    '2012.03.27

End Type

Private P_SSHIJI_O_Speck    As P_SSHIJI_O_FSpeck
Private Function P_SSHIJI_O_Create() As Integer
'********************************************************************
'*
'*              商品化指図(親)ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*      2007.11.14  :KEY3(発行日+取引先区分+手配先コード)　追加
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SSHIJI_O_Create = True
                                            'コードマスタフルパス取込み
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SSHIJI_O_Speck.fs.recoleng = Len(P_SSHIJI_O_REC)  ' レコード長
    P_SSHIJI_O_Speck.fs.PageSize = P_SSHIJI_O_PG_SIZ    ' ページサイズ
    P_SSHIJI_O_Speck.fs.idexnumb = 7                    ' インデックス数
    P_SSHIJI_O_Speck.fs.fileflag = 0                    ' ファイルフラグ
    P_SSHIJI_O_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
'2007.11.28    P_SSHIJI_O_Speck.ks0.keypos = 1              ' キーポジション
'2007.11.28    P_SSHIJI_O_Speck.ks0.keyleng = 5             ' キー長
    
    
    P_SSHIJI_O_Speck.ks0.keypos = 460                   ' キーポジション    2007.11.28
    P_SSHIJI_O_Speck.ks0.keyleng = 8                    ' キー長            2007.11.28
    
    P_SSHIJI_O_Speck.ks0.keyflag = BtKfExt              ' キーフラグ
    P_SSHIJI_O_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks0.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー０ △
    
    
    '--------------------------------------------------- キー１ ▽
    P_SSHIJI_O_Speck.ks1.keypos = 267                   ' キーポジション
    P_SSHIJI_O_Speck.ks1.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks1.reserve = &H0                  ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks2.keypos = 38                    ' キーポジション
    P_SSHIJI_O_Speck.ks2.keyleng = 2                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks2.reserve = &H0                  ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks3.keypos = 40                    ' キーポジション
    P_SSHIJI_O_Speck.ks3.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks3.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks4.keypos = 41                    ' キーポジション
    P_SSHIJI_O_Speck.ks4.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks4.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks5.keypos = 42                    ' キーポジション
    P_SSHIJI_O_Speck.ks5.keyleng = 20                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks5.reserve = &H0                  ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks6.keypos = 268                   ' キーポジション
    P_SSHIJI_O_Speck.ks6.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SSHIJI_O_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks6.reserve = &H0                  ' 予約済み
    
    
    
'2007.11.28    P_SSHIJI_O_Speck.ks7.keypos = 1                     ' キーポジション
'2007.11.28    P_SSHIJI_O_Speck.ks7.keyleng = 5                    ' キー長
                                                        
    P_SSHIJI_O_Speck.ks7.keypos = 460                   ' キーポジション    2007.11.28
    P_SSHIJI_O_Speck.ks7.keyleng = 8                    ' キー長            2007.11.28
                                                        
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks7.keyflag = BtKfExt + BtKfChg
    P_SSHIJI_O_Speck.ks7.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks7.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２  ▽
    P_SSHIJI_O_Speck.ks8.keypos = 452                   ' キーポジション
    P_SSHIJI_O_Speck.ks8.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks8.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks8.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー２ △
    
    
    
    '--------------------------------------------------- キー３ ▽
    P_SSHIJI_O_Speck.ks9.keypos = 6                     ' キーポジション
    P_SSHIJI_O_Speck.ks9.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks9.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks9.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks10.keypos = 142                  ' キーポジション
    P_SSHIJI_O_Speck.ks10.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks10.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks10.reserve = &H0                 ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks11.keypos = 73                   ' キーポジション
    P_SSHIJI_O_Speck.ks11.keyleng = 5                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks11.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks11.reserve = &H0                 ' 予約済み
    '--------------------------------------------------- キー３ △
    
    
    
    '--------------------------------------------------- キー４ ▽  2011.11.11
    P_SSHIJI_O_Speck.ks12.keypos = 38                   ' キーポジション
    P_SSHIJI_O_Speck.ks12.keyleng = 2                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks12.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks12.reserve = &H0                 ' 予約済み
    
    P_SSHIJI_O_Speck.ks13.keypos = 40                   ' キーポジション
    P_SSHIJI_O_Speck.ks13.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks13.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks13.reserve = &H0                 ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks14.keypos = 41                   ' キーポジション
    P_SSHIJI_O_Speck.ks14.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks14.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks14.reserve = &H0                 ' 予約済み
    
    
    P_SSHIJI_O_Speck.ks15.keypos = 42                   ' キーポジション
    P_SSHIJI_O_Speck.ks15.keyleng = 20                  ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks15.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks15.reserve = &H0                 ' 予約済み
    
    P_SSHIJI_O_Speck.ks16.keypos = 14                   ' キーポジション
    P_SSHIJI_O_Speck.ks16.keyleng = 14                  ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks16.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks16.reserve = &H0                 ' 予約済み
    
    
    '--------------------------------------------------- キー４ △ 2011.11.11
    
    '--------------------------------------------------- キー５  ▽ 2012.03.08
    P_SSHIJI_O_Speck.ks17.keypos = 38                   ' キーポジション
    P_SSHIJI_O_Speck.ks17.keyleng = 2                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks17.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks17.reserve = &H0                 ' 予約済み
    
    P_SSHIJI_O_Speck.ks18.keypos = 40                   ' キーポジション
    P_SSHIJI_O_Speck.ks18.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks18.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks18.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks19.keypos = 41                   ' キーポジション
    P_SSHIJI_O_Speck.ks19.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks19.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks19.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks20.keypos = 42                   ' キーポジション
    P_SSHIJI_O_Speck.ks20.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks20.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks20.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks20.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_O_Speck.ks21.keypos = 267                   ' キーポジション
    P_SSHIJI_O_Speck.ks21.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks21.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_O_Speck.ks21.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー５ △ 2012.03.08
    
    
    '--------------------------------------------------- キー６  ▽ 2012.03.27
    
    P_SSHIJI_O_Speck.ks22.keypos = 452                  ' キーポジション
    P_SSHIJI_O_Speck.ks22.keyleng = 8                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SSHIJI_O_Speck.ks22.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks22.reserve = &H0                 ' 予約済み
    
    P_SSHIJI_O_Speck.ks23.keypos = 493                  ' キーポジション
    P_SSHIJI_O_Speck.ks23.keyleng = 2                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_O_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SSHIJI_O_Speck.ks23.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_O_Speck.ks23.reserve = &H0                 ' 予約済み
    
    '--------------------------------------------------- キー６ △ 2012.03.27
    
    
    
    
    
    
    sts = BTRV(BtOpCreate, P_SSHIJI_O_POS, P_SSHIJI_O_Speck, Len(P_SSHIJI_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化指図(親)ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SSHIJI_O_Create = False

End Function

Public Function P_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図(親)ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SSHIJI_O_Open = True
                                            '商品化指図(親)ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SSHIJI_O_Create()   '商品化指図(親)ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SSHIJI_O_Open = False

End Function

