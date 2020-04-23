Attribute VB_Name = "P_Global"
Option Explicit

'********************************************************************
'*                            変数定義                              *
'*                                                                  *
'********************************************************************



'------------------------------------------ 管理マスタKEY定義
Public Const P_ST_KANRI_No$ = "01"          'ｽﾀﾝﾀﾞｰﾄKEY

Public Const P_ST_KANRI_DEF_No$ = "02"      '初期値KEY

'------------------------------------------ コマンドボタン定義
Public Const P_CMD_Upd% = 0                 '更新
Public Const P_CMD_Ins% = 2                 '新規
Public Const P_CMD_DEL% = 3                 '削除
Public Const P_CMD_DSP% = 4                 '検索/表示
Public Const P_CMD_OUT% = 7                 'ﾃﾞｰﾀ出力
Public Const P_CMD_PRT% = 8                 '印刷

Public Const P_CMD_End% = 11                '終了

'------------------------------------------ コードマスタ区分定義
Public Const P_KBN01_CD$ = "01"             '仕入区分　     コード
Public Const P_KBN01_NM$ = "仕入区分"       '          名称
Public Const P_KBN01_Len% = 2               '          有効桁数
Public Const P_KBN01_OP1% = True            '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN01_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN01_OP_NM1$ = "集計先"     '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN01_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2



Public Const P_KBN02_CD$ = "02"             '販売区分　     コード
Public Const P_KBN02_NM$ = "販売区分"       '          名称
Public Const P_KBN02_Len% = 2               '          有効桁数
Public Const P_KBN02_OP1% = True            '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN02_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN02_OP_NM1$ = "集計先"     '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN02_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2


Public Const P_KBN03_CD$ = "03"             '収支単位　     コード
Public Const P_KBN03_NM$ = "収支単位"       '          名称
Public Const P_KBN03_Len% = 3               '          有効桁数
Public Const P_KBN03_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN03_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN03_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN03_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2


Public Const P_KBN04_CD$ = "04"             '仕向け先　     コード
Public Const P_KBN04_NM$ = "仕向け先"       '          名称
Public Const P_KBN04_Len% = 2               '          有効桁数
Public Const P_KBN04_OP1% = True            '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN04_OP2% = True            '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN04_OP_NM1$ = "事業部"     '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN04_OP_NM2$ = "国内外"     '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN05_CD$ = "05"             '収単/担当者    コード
Public Const P_KBN05_NM$ = "収単／担当者"   '          名称
Public Const P_KBN05_Len% = 2               '          有効桁数
Public Const P_KBN05_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN05_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN05_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN05_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN06_CD$ = "06"             '分類           コード
Public Const P_KBN06_NM$ = "種別"           '          名称
Public Const P_KBN06_Len% = 2               '          有効桁数
Public Const P_KBN06_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN06_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN06_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN06_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN07_CD$ = "07"             '会社/事業部    コード
Public Const P_KBN07_NM$ = "会社/事業部"    '          名称
Public Const P_KBN07_Len% = 2               '          有効桁数
Public Const P_KBN07_OP1% = True            '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN07_OP2% = True            '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN07_OP_NM1$ = "事業部"     '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN07_OP_NM2$ = "国内外"     '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN08_CD$ = "08"             '資材区分       コード
Public Const P_KBN08_NM$ = "資材区分"       '          名称
Public Const P_KBN08_Len% = 1               '          有効桁数
Public Const P_KBN08_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN08_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN08_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN08_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN09_CD$ = "09"             '経営項目       コード      2008.02.28
Public Const P_KBN09_NM$ = "経営項目"       '          名称
Public Const P_KBN09_Len% = 2               '          有効桁数
Public Const P_KBN09_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN09_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN09_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN09_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2

Public Const P_KBN10_CD$ = "10"             '部署           コード      2008.02.28
Public Const P_KBN10_NM$ = "部署"           '          名称
Public Const P_KBN10_Len% = 2               '          有効桁数
Public Const P_KBN10_OP1% = False           '          ｵﾌﾟｼｮﾝ1
Public Const P_KBN10_OP2% = False           '          ｵﾌﾟｼｮﾝ2
Public Const P_KBN10_OP_NM1$ = ""           '          ｵﾌﾟｼｮﾝ名称1
Public Const P_KBN10_OP_NM2$ = ""           '          ｵﾌﾟｼｮﾝ名称2





Public Const P_KBN_MAX% = 9                 '区分数（実数-１）

Public G_SCREEN_FLG As Integer              '画面遷移用の共通フラグ

Public Const G_SCREEN_INS% = 1              '対象レコードなし
Public Const G_SCREEN_UPD% = 2              '対象レコードあり

Public Const G_INPUT_OK& = &H80000005       '入力OKﾌｨｰﾙﾄﾞ
Public Const G_INPUT_NG& = &H8000000F       '入力NGﾌｨｰﾙﾄﾞ


Public P_YOIN_TU_NYUKA      As String * 2   '「資材通常入荷」の要因
Public P_YOIN_MAE_SOUSAI    As String * 2   '「資材前借り相殺」の要因



'------------------------------------------ 組立製品
Public Const P_ASSEMBLY_OFF$ = "0"          '組立てなし
Public Const P_ASSEMBLY_ON$ = "1"           '組立てあり
'------------------------------------------ 紙
Public Const L_PAPER_OFF$ = "0"             'OFF
Public Const L_PAPER_ON$ = "1"              'ON
'------------------------------------------ プラスチック
Public Const L_PLASTIC_OFF$ = "0"           'OFF
Public Const L_PLASTIC_ON$ = "1"            'ON
'------------------------------------------ 適用機種ラベル
Public Const L_LABEL_OFF$ = "0"             'OFF
Public Const L_LABEL_ON$ = "1"              'ON
'------------------------------------------ 枚数ラベル
Public Const L_MAISU_OFF$ = "0"             'OFF
Public Const L_MAISU_ON$ = "1"              'ON
'------------------------------------------ 個装/外装/同梱・構成
Public Const P_HEAD$ = "0"                  'ﾍｯﾀﾞｰ

Public Const P_KOSOU$ = "1"                 '個装資材
Public Const P_GAISOU$ = "2"                '外装資材
Public Const P_DOUKON$ = "3"                '同梱・構成
'------------------------------------------ 取引先区分
Public Const P_TORI_GENERAL$ = "0"          '一般
Public Const P_TORI_NAISYOKU$ = "1"         '内職
Public Const P_TORI_GENKIN$ = "2"           '現金
Public Const P_TORI_SYANAI$ = "3"           '自ｾﾝﾀｰ
Public Const P_TORI_ANOTHER$ = "4"          '他ｾﾝﾀｰ
Public Const P_TORI_JIKYU$ = "5"            '内職(時給)


Public Const P_TORI_GENERAL_N$ = "一　般"
Public Const P_TORI_NAISYOKU_N$ = "内　職"
Public Const P_TORI_GENKIN_N$ = "現　金"
Public Const P_TORI_SYANAI_N$ = "自ｾﾝﾀｰ"
Public Const P_TORI_ANOTHER_N$ = "他ｾﾝﾀｰ"
Public Const P_TORI_JIKYU_N$ = "時　給"
'------------------------------------------ ﾗﾍﾞﾙ貼り計上なし
Public Const P_G_LABEL_OFF$ = "0"           '計上しない
Public Const P_G_LABEL_ON$ = "1"            '計上する
'------------------------------------------ 在庫管理
Public Const P_ZAIKO_F_ON$ = "1"            '対象
Public Const P_ZAIKO_F_OFF$ = "0"           '対象外

'------------------------------------------ 見本作成
Public Const P_SAMPLE_F_OFF$ = "0"          'なし
Public Const P_SAMPLE_F_ON$ = "1"           'あり
'------------------------------------------ 指示形態
Public Const P_SHIJI_F_NORMAL$ = "0"        'なし
Public Const P_SHIJI_F_SPOT$ = "1"          'ｽﾎﾟｯﾄ
Public Const P_SHIJI_F_KEPPIN$ = "2"        '欠品解除

Public Const P_SHIJI_F_SAIKON$ = "3"        '再梱包 2007.11.09


'------------------------------------------ 出力対象　指図票
Public Const P_PRI_SHIJI_OFF$ = "0"         'なし
Public Const P_PRI_SHIJI_ON$ = "1"          'あり
'------------------------------------------ 出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ
Public Const P_PRI_PARTS_OFF$ = "0"         'なし
Public Const P_PRI_PARTS_ON$ = "1"          'あり
'------------------------------------------ 出力対象　外装ﾗﾍﾞﾙ
Public Const P_PRI_GAISOU_OFF$ = "0"        'なし
Public Const P_PRI_GAISOU_ON$ = "1"         'あり
'------------------------------------------ 出力対象　機種ﾗﾍﾞﾙ
Public Const P_PRI_KISHU_OFF$ = "0"        'なし
Public Const P_PRI_KISHU_ON$ = "1"         'あり

'------------------------------------------ 完了F
Public Const P_KAN_OFF$ = "0"               '未完
Public Const P_KAN_ON$ = "1"                '完了

'------------------------------------------ ｷｬﾝｾﾙF
Public Const P_CANCEL_OFF$ = "0"            '未
Public Const P_CANCEL_ON$ = "1"             'ｷｬﾝｾﾙ

'------------------------------------------ 受入F
Public Const P_UKEIRE_CON$ = "0"            '継続（未完）
Public Const P_UKEIRE_END$ = "1"            '最終受入

'------------------------------------------ 印刷F
Public Const P_PRINT_OFF$ = "0"             '未印刷
Public Const P_PRINT_ON$ = "1"              '印刷済

'------------------------------------------ 請求F
Public Const P_SEIKYU_NON$ = "0"            '未処理
Public Const P_SEIKYU_PRI$ = "1"            '印刷済
Public Const P_SEIKYU_END$ = "9"            '締め済


'------------------------------------------ 販売区分
Public Const P_HN_HANBAI$ = "1"             '販売
Public Const P_HN_SEIZOU$ = "2"             '製造
Public Const P_HN_YATIN$ = "3"              '家賃
Public Const P_HN_ETC$ = "4"                'その他
Public Const P_HN_HAKEN$ = "5"              '派遣
                                            
                                            '*上記以外は全てその他へ
'------------------------------------------ 仕入区分
Public Const P_SH_SHIIRE$ = "1"             '仕入
Public Const P_SH_SEIZOU$ = "2"             '製造
Public Const P_SH_YATIN$ = "3"              '家賃
Public Const P_SH_ETC$ = "4"                'その他
Public Const P_SH_HAKEN$ = "5"              '派遣
Public Const P_SH_KEIHI$ = "6"              '経費
Public Const P_SH_ZEI$ = "7"                '消費税


'------------------------------------------ 棚卸しﾃﾞｰﾀ合計KEY(仕向け先)
Public Const P_StokSum_Key$ = "zzz"
'------------------------------------------ 生産実績合計KEY(ｸﾗｽ)
Public Const P_ClassSum_Key$ = "!!!!!!!!!!!!!!!!!!!!"


