Attribute VB_Name = "O_Y_NYU"
Option Explicit
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ファイル定義                        *
'*                                                                  *
'********************************************************************
'ファイルＩＤ
Public Const O_Y_NYU_ID$ = "O_Y_NYU"

'ページサイズ
Public Const O_Y_NYU_PG_SIZ% = 2048

'ポジション・ブロック
Public O_Y_NYU_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type O_Y_NYUREC_Tag
    KAN_KBN(0 To 0)             As Byte     '完了区分
    DT_SYU(0 To 0)              As Byte     'データ種別
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    TEXT_NO(0 To 8)             As Byte     'テキスト№
    JGYOBA(0 To 7)              As Byte     '事業場
    DATA_KBN(0 To 0)            As Byte     'データ区分
    TORI_KBN(0 To 1)            As Byte     '取引区分
    ID_NO(0 To 7)               As Byte     'ID-NO
    HIN_NO(0 To 19)             As Byte     '品目番号
    DEN_NO(0 To 9)              As Byte     '伝票番号
    SURYO(0 To 6)               As Byte     '出庫数量
    MUKE_CODE(0 To 7)           As Byte     '出庫先
    SYUKO_SYUSI(0 To 1)         As Byte     '出庫収支
    SYUKO_YMD(0 To 7)           As Byte     '出庫日付
    TANKA(0 To 9)               As Byte     '単価
    ODER_NO(0 To 11)            As Byte     'オーダー番号
    ITEM_NO(0 To 4)             As Byte     'アイテム番号
    ODER_R_NO(0 To 4)           As Byte     'オーダー略号
    KOSO_KEITAI(0 To 9)         As Byte     '個装形態
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TANABAN1(0 To 9)            As Byte     '棚番１
    TANABAN2(0 To 9)            As Byte     '棚番２
    TANABAN3(0 To 9)            As Byte     '棚番３
    MUKE_NAME(0 To 23)          As Byte     '出庫先名称
    CYU_KBN(0 To 0)             As Byte     '注文区分
    CYU_KBN_NAME(0 To 9)        As Byte     '注文区分名称
    ORIGIN1(0 To 9)             As Byte     '原産国１
    ORIGIN2(0 To 9)             As Byte     '原産国２
    BIKOU2(0 To 39)             As Byte     '備考２
    HAN_KBN(0 To 0)             As Byte     '販売区分
    CHOKU_KBN(0 To 0)           As Byte     '直送区分
    UNIT_ID_NO(0 To 7)          As Byte     'ﾕﾆｯﾄ修理ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte     '在庫引当順序
    GOKON_KANRI_NO(0 To 8)      As Byte     '合梱管理番号
    JUCHU_ZAN(0 To 6)           As Byte     '受注残数量
    KYOKYU_KBN(0 To 0)          As Byte     '供給区分
    SHOHIN_SYUSI(0 To 1)        As Byte     '商品化納入先収支
    BIKOU1(0 To 39)             As Byte     '備考１
    CHOHA_KBN(0 To 0)           As Byte     '帳端区分
    JYU_HIN_NO(0 To 19)         As Byte     '受注品目番号
    HIN_NAME(0 To 19)           As Byte     '品名
    HIN_CHANGE_KBN(0 To 0)      As Byte     '品番変更区分
    MODULE_EXCHANGE(0 To 0)     As Byte     'モジュール交換区分
    ZAIKO_SYUSI(0 To 1)         As Byte     '残在庫まとめ在庫収支コード
    NOUKI_YMD(0 To 7)           As Byte     '指定納期
    SERVICE_KANRI_NO(0 To 8)    As Byte     'サービス会社管理番号
    KI_HIN_NO(0 To 2)           As Byte     '機種品目コード
    ENVIRONMENT_KBN(0 To 0)     As Byte     '環境規格部品区分
    KAN_DT(0 To 7)              As Byte     '完了日付
    BEF_NYU_QTY(0 To 7)         As Byte     '先行入荷数
    YOSAN_FROM(0 To 4)          As Byte     '予算単位（元）
    YOSAN_TO(0 To 4)            As Byte     '予算単位（先）
    HTANABAN(0 To 7)            As Byte     '標準棚番
    HIN_NAI(0 To 12)            As Byte     '品番（内部）
    FILLER(0 To 64)             As Byte
End Type

'データ・バッファ
Public O_Y_NYUREC                  As O_Y_NYUREC_Tag

'キー定義
Type KEY0_O_Y_NYU            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TEXT_NO(0 To 8)             As Byte     'テキスト№
End Type

Type KEY1_O_Y_NYU            'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部区分
    KAN_KBN(0 To 0)             As Byte     '完了区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品目番号
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    TEXT_NO(0 To 8)             As Byte     'テキスト№
End Type

Type KEY2_O_Y_NYU            'ＫＥＹ２
    JGYOBU(0 To 0)              As Byte     '事業部区分
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
    HIN_NO(0 To 19)             As Byte     '品目番号
    NAIGAI(0 To 0)              As Byte     '国内外
End Type

Type KEY3_O_Y_NYU            'ＫＥＹ３
    SYUKA_YMD(0 To 7)           As Byte     '出荷日
End Type



'キー・データ
Public K0_O_Y_NYU                 As KEY0_O_Y_NYU
Public K1_O_Y_NYU                 As KEY1_O_Y_NYU
Public K2_O_Y_NYU                 As KEY2_O_Y_NYU
Public K3_O_Y_NYU                 As KEY3_O_Y_NYU

Private Type O_Y_NYU_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
    ks11    As BtKeySpeck
    ks12    As BtKeySpeck
    ks13    As BtKeySpeck
    ks14    As BtKeySpeck
End Type

Private O_Y_NYU_Speck As O_Y_NYU_FSpeck

Private Function O_Y_NYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    O_Y_NYU_Create = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", O_Y_NYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_NYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_Y_NYU_Speck.fs.recoleng = Len(O_Y_NYUREC)     ' レコード長
    O_Y_NYU_Speck.fs.PageSize = O_Y_NYU_PG_SIZ      ' ページサイズ
    O_Y_NYU_Speck.fs.idexnumb = 4                 ' インデックス数
    O_Y_NYU_Speck.fs.fileflag = 0                 ' ファイルフラグ
    O_Y_NYU_Speck.fs.reserve = &H0                ' 予約済み
    '-------------------------------------------
                                                ' キー０
    O_Y_NYU_Speck.ks0.keypos = 3                  ' キーポジション
    O_Y_NYU_Speck.ks0.keyleng = 1                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks0.reserve = &H0               ' 予約済み
                                                ' キー０
    O_Y_NYU_Speck.ks1.keypos = 130                ' キーポジション
    O_Y_NYU_Speck.ks1.keyleng = 8                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks1.reserve = &H0               ' 予約済み
                                                ' キー０
    O_Y_NYU_Speck.ks2.keypos = 5                  ' キーポジション
    O_Y_NYU_Speck.ks2.keyleng = 9                 ' キー長
    O_Y_NYU_Speck.ks2.keyflag = BtKfExt           ' キーフラグ
    O_Y_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks2.reserve = &H0               ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー１
    O_Y_NYU_Speck.ks3.keypos = 3                  ' キーポジション
    O_Y_NYU_Speck.ks3.keyleng = 1                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks3.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks3.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks3.reserve = &H0               ' 予約済み
                                                ' キー１
    O_Y_NYU_Speck.ks4.keypos = 1                  ' キーポジション
    O_Y_NYU_Speck.ks4.keyleng = 1                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks4.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks4.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks4.reserve = &H0               ' 予約済み
                                                ' キー１
    O_Y_NYU_Speck.ks5.keypos = 4                 ' キーポジション
    O_Y_NYU_Speck.ks5.keyleng = 1                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks5.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks5.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks5.reserve = &H0               ' 予約済み
                                                ' キー１
    O_Y_NYU_Speck.ks6.keypos = 33                 ' キーポジション
    O_Y_NYU_Speck.ks6.keyleng = 20                ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks6.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks6.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks6.reserve = &H0               ' 予約済み
                                                ' キー１
    O_Y_NYU_Speck.ks7.keypos = 130                ' キーポジション
    O_Y_NYU_Speck.ks7.keyleng = 8                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks7.keyflag = BtKfExt + BtKfSeg + BtKfChg
    O_Y_NYU_Speck.ks7.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks7.reserve = &H0               ' 予約済み
                                                ' キー１
    O_Y_NYU_Speck.ks8.keypos = 5                ' キーポジション
    O_Y_NYU_Speck.ks8.keyleng = 9                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks8.keyflag = BtKfExt + BtKfChg
    O_Y_NYU_Speck.ks8.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks8.reserve = &H0               ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー２
    O_Y_NYU_Speck.ks9.keypos = 3                  ' キーポジション
    O_Y_NYU_Speck.ks9.keyleng = 1                 ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks9.keytype = Chr(BtKtString)   ' キータイプ
    O_Y_NYU_Speck.ks9.reserve = &H0               ' 予約済み
                                                ' キー２
    O_Y_NYU_Speck.ks10.keypos = 130               ' キーポジション
    O_Y_NYU_Speck.ks10.keyleng = 8                ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks10.keytype = Chr(BtKtString)  ' キータイプ
    O_Y_NYU_Speck.ks10.reserve = &H0              ' 予約済み
                                                ' キー２
    O_Y_NYU_Speck.ks11.keypos = 33                ' キーポジション
    O_Y_NYU_Speck.ks11.keyleng = 20               ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks11.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks11.keytype = Chr(BtKtString)  ' キータイプ
    O_Y_NYU_Speck.ks11.reserve = &H0              ' 予約済み
                                                ' キー２
    O_Y_NYU_Speck.ks12.keypos = 4                 ' キーポジション
    O_Y_NYU_Speck.ks12.keyleng = 1                ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks12.keyflag = BtKfExt + BtKfSeg
    O_Y_NYU_Speck.ks12.keytype = Chr(BtKtString)  ' キータイプ
    O_Y_NYU_Speck.ks12.reserve = &H0              ' 予約済み
                                                ' キー２
    O_Y_NYU_Speck.ks13.keypos = 5                 ' キーポジション
    O_Y_NYU_Speck.ks13.keyleng = 9                ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks13.keyflag = BtKfExt
    O_Y_NYU_Speck.ks13.keytype = Chr(BtKtString)  ' キータイプ
    O_Y_NYU_Speck.ks13.reserve = &H0              ' 予約済み
    '-------------------------------------------
                                                
                                                ' キー３
    O_Y_NYU_Speck.ks14.keypos = 130                ' キーポジション
    O_Y_NYU_Speck.ks14.keyleng = 8                ' キー長
                                                ' キーフラグ
    O_Y_NYU_Speck.ks14.keyflag = BtKfExt + BtKfDup
    O_Y_NYU_Speck.ks14.keytype = Chr(BtKtString)  ' キータイプ
    O_Y_NYU_Speck.ks14.reserve = &H0              ' 予約済み
    '-------------------------------------------
    
    sts = BTRV(BtOpCreate, O_Y_NYU_POS, O_Y_NYU_Speck, Len(O_Y_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入荷予定データ")
        O_Y_NYU_Create = True
        Exit Function
    End If

    O_Y_NYU_Create = False

End Function

Function O_Y_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ  ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    O_Y_NYU_Open = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", O_Y_NYU_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_Y_NYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_Y_NYU_Create()        '入荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入荷予定データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "入荷予定データ")
                Exit Function
        End Select
    Loop
    
    O_Y_NYU_Open = False

End Function


