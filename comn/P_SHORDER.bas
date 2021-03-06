Attribute VB_Name = "P_SHORDER"
Option Explicit

'********************************************************************
'*
'*              資材注文ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SHORDER_ID$ = "P_SHORDER"

'ページサイズ
Private Const P_SHORDER_PG_SIZ% = 4096

'ポジション・ブロック
Public P_SHORDER_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_SHORDER_REC_Tag
    
    ORDER_NO(0 To 4)        As Byte         '注文��
    ORDER_DT(0 To 7)        As Byte         '注文日
    Print_datetime(0 To 13) As Byte         '発行日時
    TANTO_CODE(0 To 4)      As Byte         '担当者ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    DELI_CODE(0 To 4)       As Byte         '納入先ｺｰﾄﾞ
    ORDER_QTY(0 To 10)      As Byte         '注文数(9(8)V99)
    Y_NOUKI_DT(0 To 7)      As Byte         '予定納期
    TANKA(0 To 10)          As Byte         '発注単価(9(8)V99)
    LOT(0 To 7)             As Byte         '発注ﾛｯﾄ
    KAN_F(0 To 0)           As Byte         '完了F
    KAN_DT(0 To 7)          As Byte         '完了日
    BUNNOU_CNT(0 To 1)      As Byte         '分納回数
    UKEIRE_QTY(0 To 10)     As Byte         '受入数（合計）(9(8)V99)
    
    CANCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
    PRINT_F(0 To 0)         As Byte         '注文書印刷ﾌﾗｸﾞ
    WS_NO(0 To 9)           As Byte         '入力端末
    G_SHIIRE_KBN(0 To 1)    As Byte         '仕入区分
    G_SYUSHI(0 To 2)        As Byte         '収支単位
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    
    ANS_NOUKI_DT(0 To 7)    As Byte         '納期回答日   2008.01.10
    USE_YM(0 To 5)          As Byte         '使用年月     2008.01.10
    
    
    UPD_FLG(0 To 0)         As Byte         '展開更新済み   2012.01.17
    
    FILLER(0 To 70)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SHORDER_REC       As P_SHORDER_REC_Tag

'キー定義

Public Type KEY0_P_SHORDER                         'ＫＥＹ０
    ORDER_NO(0 To 4)        As Byte         '注文��
End Type
    
Public Type KEY1_P_SHORDER                         'ＫＥＹ１
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    ORDER_DT(0 To 7)        As Byte         '注文日
    ORDER_NO(0 To 4)        As Byte         '注文��
End Type
    
Public Type KEY2_P_SHORDER                         'ＫＥＹ２
    WS_NO(0 To 9)           As Byte         '入力端末
    PRINT_F(0 To 0)         As Byte         '注文書印刷ﾌﾗｸﾞ
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    ORDER_NO(0 To 4)        As Byte         '注文��
End Type
    
Public Type KEY3_P_SHORDER                         'ＫＥＹ３
    KAN_F(0 To 0)           As Byte         '完了F
    ORDER_DT(0 To 7)        As Byte         '注文日
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
End Type
    
    
Public Type KEY4_P_SHORDER                         'ＫＥＹ４
    KAN_F(0 To 0)           As Byte         '完了F
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    ORDER_DT(0 To 7)        As Byte         '注文日
End Type
    
Public Type KEY5_P_SHORDER                         'ＫＥＹ５    2007.12.05
    KAN_F(0 To 0)           As Byte         '完了F
    Y_NOUKI_DT(0 To 7)      As Byte         '予定納期
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
End Type
    
    
Public Type KEY6_P_SHORDER                         'ＫＥＹ６    2008.03.23
    ANS_NOUKI_DT(0 To 7)    As Byte         '納期回答日
End Type
    
    
Public Type KEY7_P_SHORDER                         'ＫＥＹ７    2012.03.06
    USE_YM(0 To 5)          As Byte         '使用年月     2008.01.10
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    CANCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
End Type
    
Public Type KEY8_P_SHORDER                         'ＫＥＹ８    2019.03.15
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番

    KAN_F(0 To 0)           As Byte         '完了F

    CANCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF

    Print_datetime(0 To 13) As Byte         '発行日時

End Type
    
    
    
    
    
'キー・データ
Public K0_P_SHORDER         As KEY0_P_SHORDER
Public K1_P_SHORDER         As KEY1_P_SHORDER
Public K2_P_SHORDER         As KEY2_P_SHORDER
Public K3_P_SHORDER         As KEY3_P_SHORDER
Public K4_P_SHORDER         As KEY4_P_SHORDER
Public K5_P_SHORDER         As KEY5_P_SHORDER   '2007.12.05

Public K6_P_SHORDER         As KEY6_P_SHORDER   '2008.03.23

Public K7_P_SHORDER         As KEY7_P_SHORDER   '2012.03.06

Public K8_P_SHORDER         As KEY8_P_SHORDER   '2019.03.15



Type P_SHORDER_FSpeck
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
    ks12                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks13                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks14                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks15                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks16                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2007.12.05
    ks17                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2007.12.05
    ks18                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2007.12.05

    ks19                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2008.03.23

    ks20                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.06
    ks21                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.06
    ks22                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.06
    ks23                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.06
    ks24                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.06

    ks25                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15
    ks26                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15
    ks27                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15
    ks28                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15
    ks29                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15
    ks30                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2019.03.15

End Type

Private P_SHORDER_Speck    As P_SHORDER_FSpeck
Private Function P_SHORDER_Create() As Integer
'********************************************************************
'*
'*              資材注文ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHORDER_Create = True
                                            '資材注文ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHORDER_Speck.fs.recoleng = Len(P_SHORDER_REC)    ' レコード長
    P_SHORDER_Speck.fs.PageSize = P_SHORDER_PG_SIZ      ' ページサイズ
    P_SHORDER_Speck.fs.idexnumb = 9                     ' インデックス数
    P_SHORDER_Speck.fs.fileflag = 0                     ' ファイルフラグ
    P_SHORDER_Speck.fs.reserve = &H0                    ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHORDER_Speck.ks0.keypos = 1                      ' キーポジション
    P_SHORDER_Speck.ks0.keyleng = 5                     ' キー長
    P_SHORDER_Speck.ks0.keyflag = BtKfExt               ' キーフラグ
    P_SHORDER_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks0.reserve = &H0                   ' 予約済み
    
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SHORDER_Speck.ks1.keypos = 33                     ' キーポジション
    P_SHORDER_Speck.ks1.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks1.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks2.keypos = 34                     ' キーポジション
    P_SHORDER_Speck.ks2.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks2.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks3.keypos = 35                     ' キーポジション
    P_SHORDER_Speck.ks3.keyleng = 20                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks3.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks4.keypos = 6                      ' キーポジション
    P_SHORDER_Speck.ks4.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks4.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks5.keypos = 1                      ' キーポジション
    P_SHORDER_Speck.ks5.keyleng = 5                     ' キー長
    P_SHORDER_Speck.ks5.keyflag = BtKfExt + BtKfChg     ' キーフラグ
    P_SHORDER_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks5.reserve = &H0                   ' 予約済み
    
    '--------------------------------------------------- キー１ △
    
    
    
    '--------------------------------------------------- キー２ ▽
    P_SHORDER_Speck.ks6.keypos = 141                    ' キーポジション
    P_SHORDER_Speck.ks6.keyleng = 10                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks6.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks7.keypos = 140                    ' キーポジション
    P_SHORDER_Speck.ks7.keyleng = 1                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks7.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks7.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks8.keypos = 55                     ' キーポジション
    P_SHORDER_Speck.ks8.keyleng = 5                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfSeg
    P_SHORDER_Speck.ks8.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks8.reserve = &H0                   ' 予約済み
    
    P_SHORDER_Speck.ks9.keypos = 1                      ' キーポジション
    P_SHORDER_Speck.ks9.keyleng = 5                     ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks9.keyflag = BtKfExt + BtKfChg
    P_SHORDER_Speck.ks9.keytype = Chr(BtKtString)       ' キータイプ
    P_SHORDER_Speck.ks9.reserve = &H0                   ' 予約済み
    
    '--------------------------------------------------- キー２ △
    
    '--------------------------------------------------- キー３ ▽
    
    P_SHORDER_Speck.ks10.keypos = 103                   ' キーポジション
    P_SHORDER_Speck.ks10.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks10.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks10.reserve = &H0                  ' 予約済み
    
    P_SHORDER_Speck.ks11.keypos = 6                     ' キーポジション
    P_SHORDER_Speck.ks11.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks11.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks11.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks11.reserve = &H0                  ' 予約済み
    
    
    P_SHORDER_Speck.ks12.keypos = 55                    ' キーポジション
    P_SHORDER_Speck.ks12.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks12.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks12.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks12.reserve = &H0                  ' 予約済み

    '--------------------------------------------------- キー３ △
    
    
    '--------------------------------------------------- キー４ ▽
    
    P_SHORDER_Speck.ks13.keypos = 103                   ' キーポジション
    P_SHORDER_Speck.ks13.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks13.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks13.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks13.reserve = &H0                  ' 予約済み
    
    P_SHORDER_Speck.ks14.keypos = 55                    ' キーポジション
    P_SHORDER_Speck.ks14.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks14.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks14.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks14.reserve = &H0                  ' 予約済み
    
    
    P_SHORDER_Speck.ks15.keypos = 6                    ' キーポジション
    P_SHORDER_Speck.ks15.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks15.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks15.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks15.reserve = &H0                  ' 予約済み

    '--------------------------------------------------- キー４ △
    
    
    
    
    
    '--------------------------------------------------- キー５ 2007.12.05 ▽
    
    P_SHORDER_Speck.ks16.keypos = 103                   ' キーポジション
    P_SHORDER_Speck.ks16.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks16.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks16.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks16.reserve = &H0                  ' 予約済み
    
    P_SHORDER_Speck.ks17.keypos = 76                    ' キーポジション
    P_SHORDER_Speck.ks17.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks17.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks17.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks17.reserve = &H0                  ' 予約済み
    
    P_SHORDER_Speck.ks18.keypos = 55                    ' キーポジション
    P_SHORDER_Speck.ks18.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks18.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks18.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks18.reserve = &H0                  ' 予約済み

    '--------------------------------------------------- キー５ 2007.12.05 △
    
    
    '--------------------------------------------------- キー６ 2008.03.23 ▽


    P_SHORDER_Speck.ks19.keypos = 157                   ' キーポジション
    P_SHORDER_Speck.ks19.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks19.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks19.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks19.reserve = &H0                  ' 予約済み



    '--------------------------------------------------- キー６ 2008.03.20 △
    
    
    '--------------------------------------------------- キー７ 2012.03.06 ▽
    P_SHORDER_Speck.ks20.keypos = 165                   ' キーポジション
    P_SHORDER_Speck.ks20.keyleng = 6                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks20.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks20.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks20.reserve = &H0                  ' 予約済み

    P_SHORDER_Speck.ks21.keypos = 33                    ' キーポジション
    P_SHORDER_Speck.ks21.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks21.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks21.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks21.reserve = &H0                  ' 予約済み

    P_SHORDER_Speck.ks22.keypos = 34                    ' キーポジション
    P_SHORDER_Speck.ks22.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks22.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks22.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks22.reserve = &H0                  ' 予約済み

    P_SHORDER_Speck.ks23.keypos = 35                    ' キーポジション
    P_SHORDER_Speck.ks23.keyleng = 20                   ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks23.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks23.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks23.reserve = &H0                  ' 予約済み

    P_SHORDER_Speck.ks24.keypos = 125                   ' キーポジション
    P_SHORDER_Speck.ks24.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks24.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks24.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks24.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー７ 2012.03.06 △
    
    
    
    '--------------------------------------------------- キー８ 2019.03.15 ▽


    P_SHORDER_Speck.ks25.keypos = 33                   ' キーポジション
    P_SHORDER_Speck.ks25.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks25.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks25.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks25.reserve = &H0                  ' 予約済み

    P_SHORDER_Speck.ks26.keypos = 34                   ' キーポジション
    P_SHORDER_Speck.ks26.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks26.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks26.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks26.reserve = &H0                  ' 予約済み


    P_SHORDER_Speck.ks27.keypos = 35                   ' キーポジション
    P_SHORDER_Speck.ks27.keyleng = 20                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks27.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks27.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks27.reserve = &H0                  ' 予約済み


    P_SHORDER_Speck.ks28.keypos = 103                   ' キーポジション
    P_SHORDER_Speck.ks28.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks28.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks28.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks28.reserve = &H0                  ' 予約済み


    P_SHORDER_Speck.ks29.keypos = 125                   ' キーポジション
    P_SHORDER_Speck.ks29.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks29.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    P_SHORDER_Speck.ks29.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks29.reserve = &H0                  ' 予約済み


    P_SHORDER_Speck.ks30.keypos = 14                    ' キーポジション
    P_SHORDER_Speck.ks30.keyleng = 14                   ' キー長
                                                        ' キーフラグ
    P_SHORDER_Speck.ks30.keyflag = BtKfExt + BtKfChg + BtKfDup
    P_SHORDER_Speck.ks30.keytype = Chr(BtKtString)      ' キータイプ
    P_SHORDER_Speck.ks30.reserve = &H0                  ' 予約済み



    '--------------------------------------------------- キー８ 2019.03.15 △
    
    
    
    
    sts = BTRV(BtOpCreate, P_SHORDER_POS, P_SHORDER_Speck, Len(P_SHORDER_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材注文ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SHORDER_Create = False

End Function

Public Function P_SHORDER_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材注文ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHORDER_Open = True
                                            '資材注文ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHORDER_Create()   '資材注文ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材注文ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SHORDER_Open = False

End Function

