Attribute VB_Name = "Global"
Option Explicit

'   ウインドウズ終了要求
    Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'   ウインドウズ終了要求
    Declare Function ExitWindowsEx Lib "user32 " (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'   処理中断
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'   ＩＮＩファイル書き込み
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'   ＩＮＩファイル読み込み
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'   コンピュータ名取得
    Declare Function GetComputerNameA Lib "kernel32" _
           (ByVal IpBuffer As String, nSize As Long) As Long
    
    Declare Function GetVersion Lib "kernel32.dll" () As Long
    Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

    Public Const HWND_BROADCAST  As Long = &HFFFF&
    Public Const WM_WININICHANGE As Long = &H1A&

    Public Const EM_GETLINECOUNT As Long = &HBA     '2016.01.05

    '2019.03.29
    Public Const CB_SHOWDROPDOWN = &H14F



'   キーストローク合成関数
    Declare Sub Keybd_Event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


    Declare Function GetDeviceCaps Lib "gdi32" _
        (ByVal hDC As Long, ByVal nIndex As Long) As Long

    Public Const HORZRES = 8           '実際のスクリーンの幅（実印刷領域）
    Public Const VERTRES = 10          '実際のスクリーンの高さ
    Public Const PHYSICALWIDTH = 110   '物理的幅(実用紙サイズ）
    Public Const PHYSICALHEIGHT = 111  '物理的高さ
    Public Const PHYSICALOFFSETX = 112 '印刷可能な左方向のマージン
    Public Const PHYSICALOFFSETY = 113 '印刷可能な上方向のマージン
    
'********************************************************************
'*                            変数定義                              *
'*                                                                  *
                   
'********************************************************************


'***** システム異常 ********** 97.01.08
Public Const SYS_ERR% = -100
Public Const SYS_CANCEL% = -200

'***** システム共通定義 ******
                                    
'-----------------------------------'事業部区分
Public Const SOJIKI$ = "7"          '掃除機
Public Const DENKA$ = "D"           '電化調理
Public Const SUIHAN$ = "4"          '炊飯器
Public Const SENTAKU$ = "1"         '洗濯機（アイロン）
Public Const AIRCON$ = "A"          'エアコン           2004.12.01
Public Const REIZOU$ = "R"          '冷蔵庫             2007.05.24

Public Const SETSUBI$ = "B"         '設備   2007.03.28

Public Const SHIZAI$ = "S"          '資材   2005.11.16
Public Const BUZAI$ = "C"           '部材   2012.03.23
Public Const BLBU$ = "5"            'ﾋﾞｭｰﾃｨﾘﾋﾞﾝｸﾞ   2012.04.06
Public Const OVEN$ = "6"            '電子レンジ     2012.05.16
Public Const YUKADAN$ = "Y"         '床暖房         2013.06.06
Public Const JCS$ = "J"             'JCS            2015.01.22
Public Const SHOKUSEN$ = "2"        '食洗           2015.03.03

Public Const JGYOBU_NON$ = "0"      '事業部区分なし
                                   
'-----------------------------------'倉庫区分
Public Const BUN_JITU$ = "0"        '実倉庫
Public Const BUN_KASO$ = "1"        'システム固有
Public Const BUN_AUTO$ = "2"        '自動倉庫
                                   
Public Const SOKO_BUN0$ = "実倉庫  "
Public Const SOKO_BUN1$ = "システム"
Public Const SOKO_BUN2$ = "自動倉庫"
'-----------------------------------'国内外
Public Const NAIGAI_NON$ = "0"      'なし
Public Const NAIGAI_NAI$ = "1"      '国内
Public Const NAIGAI_GAI$ = "2"      '海外
                                   
Public Const NAIGAI0$ = "なし"
Public Const NAIGAI1$ = "国内"
Public Const NAIGAI2$ = "海外"
'-----------------------------------'倉庫／棚　使用可否
Public Const KAHI_KBN_OK$ = "0"     '使用可
Public Const KAHI_KBN_NG$ = "1"     '使用不可
                                   
Public Const KAHI_KBN0$ = "使用　可"
Public Const KAHI_KBN1$ = "使用不可"
'-----------------------------------'倉庫／棚　在庫照合 2004.02
Public Const ZAIKO_SHOGO_FLG_OK$ = "0"      '照合有
Public Const ZAIKO_SHOGO_FLG_NG$ = "1"      '照合無
                                   
Public Const ZAIKO_SHOGO0$ = "対　象"
Public Const ZAIKO_SHOGO1$ = "対象外"
'-----------------------------------'倉庫 混載区分
Public Const KONS_KBN_OK$ = "0"     '混載可
Public Const KONS_KBN_NG$ = "1"     '混載不可

Public Const KONS_KBN0$ = "混載可  "
Public Const KONS_KBN1$ = "混載不可"
'-----------------------------------'向け先　ＭＴＳ区分
Public Const MUKE_MTS$ = "1"        'ＭＴＳ
Public Const MUKE_SS$ = "2"         'ＳＳ

'-----------------------------------'出荷予定／在庫　使用可否
Public Const LOCK_OFF$ = "0"        '使用可
Public Const LOCK_ON$ = "1"         '使用中
'-----------------------------------'商品可済み／未商品の識別2004.04
Public Const GOODS_ON$ = "0"        '商品化
Public Const GOODS_OFF$ = "1"       '未商品
'-----------------------------------'出荷／入荷の完了フラグ
Public Const KAN_KBN_UN$ = "0"      '未処理
Public Const KAN_KBN_FIN$ = "9"     '処理済み
'-----------------------------------'出荷予定 完了区分
'Public Const KAN_SOFF_POFF_KOFF$ = "0"      '完了区分＝未出庫／未印刷／未検品
'Public Const KAN_SING_POFF_KOFF$ = "1"      '完了区分＝出庫中／未印刷／未検品
'Public Const KAN_SOFF_PON_KOFF$ = "2"       '完了区分＝未出庫／印刷済／未検品
'Public Const KAN_SING_PON_KOFF$ = "3"       '完了区分＝出庫中／印刷済／未検品
'Public Const KAN_SON_POFF_KOFF$ = "4"       '完了区分＝出庫済／未印刷／未検品
'Public Const KAN_SON_PON_KOFF$ = "5"        '完了区分＝出庫済／印刷済／未検品
'Public Const KAN_SON_PNON_KON$ = "6"        '完了区分＝出庫済／―／検品済
'Public Const KAN_SNO_PNO_KNO$ = "9"         '完了区分＝出庫不可／印刷不可／検品不可

'Public Const KAN_L_SOFF_POFF_KOFF$ = "A"    '完了区分＝未出庫／未印刷／未検品
'Public Const KAN_L_SING_POFF_KOFF$ = "B"    '完了区分＝出庫中／未印刷／未検品
'Public Const KAN_L_SOFF_PON_KOFF$ = "C"     '完了区分＝未出庫／印刷済／未検品
'Public Const KAN_L_SING_PON_KOFF$ = "D"     '完了区分＝出庫中／印刷済／未検品

'-----------------------------------'作業／要因の識別
Public Const ACT_ZAITEI_IN$ = "1"       '在訂（＋）
Public Const ACT_ZAITEI_OUT$ = "2"      '在訂（－）
Public Const ACT_NYUKA$ = "3"           '入荷
Public Const ACT_SYUKA_KEI$ = "4"       '出荷(出荷予定有り)
Public Const ACT_SYUKA_HYO$ = "5"       '出荷(出庫表)
Public Const ACT_SYUKA_GAI$ = "6"       '出荷(出荷予定無し)
Public Const ACT_IDO_IN$ = "7"          '移動入庫
Public Const ACT_IDO_OUT$ = "8"         '移動出庫
Public Const ACT_DENPYO_ID$ = "9"       '伝票ＩＤ   2004.02
Public Const ACT_KENPIN$ = "A"          '検品
Public Const ACT_WEL_ETC$ = "B"         'WEL専用

Public Const ACT_KENPIN_MTS$ = "C"      '向け先読み込み用
Public Const ACT_GOODS_ONFF$ = "D"      '商品←→未商品切り替え用

Public Const ACT_SPECIAL_PROCESS$ = "E" '特殊処理

Public Const ACT_KENPIN_DEN$ = "F"      '検品（大阪PC） 2006.12.07

Public Const ACT_SYUKA_HYO_OSAKA$ = "G" '出庫表出庫（大阪PC） 2007.03.16

Public Const ACT_IN_KENPIN_OSAKA$ = "H" '入庫検品（大阪PC） 2007.06.07
Public Const ACT_IN_TANA_OSAKA$ = "I"   '実棚入庫（大阪PC） 2007.06.07

Public Const ACT_FURIKAE$ = "J"         '資材振替（大阪PC） 2007.06.28


Public Const ACT_BINNO$ = "K"           '便№（移管用） 2009.03.11


Public Const ACT_KENPIN_GAI$ = "L"      '検品海外   2009.08.05


'Public Const ACT_SAI_SU$ = "M"          '才数／口数   2010.01.21


Public Const ACT_SHOUHINKA$ = "M"       '商品化   2010.09.03

Public Const ACT_LotNo$ = "N"         '床暖房　特殊   2013.06.06

Public Const ACT_MODULE$ = "O"         'モジュール   2014.06.24


Public Const ACT_DENPYO_ID2$ = "P"      '伝票ＩＤ   2015.02.21

Public Const ACT_KENPIN_Drct$ = "Q"     '直送検品   2016.10.03

Public Const ACT_BCR_PRINT$ = "R"       'バーコード印字　2017.04.10

Public Const ACT_NEW_KENPIN$ = "S"      '新検品 2018.11.05
Public Const ACT_NEW_KENPIN_MTS$ = "T"  '新向け先読み込み用 2018.11.05





Public Const ACT_SYSTEM$ = "Z"      'システム専用

Public YOIN_TU_NYUKA        As String * 2       '「通常入荷」の要因
Public YOIN_MAEGARI         As String * 2       '「前借り入荷」の要因
Public YOIN_MAE_SOUSAI      As String * 2       '「前借り相殺」の要因
Public YOIN_FURIKAE         As String * 2       '「国内外振替え」の要因
Public YOIN_FURIKAE_OUT     As String * 2       '「国内外振替え事の出庫」の要因
Public YOIN_FURIKAE_IN      As String * 2       '「国内外振替え事の入庫」の要因

Public YOIN_TANASHOGO       As String * 2       '「棚照合」の要因
Public YOIN_TANAHINSHOGO    As String * 2       '「棚品照合」の要因


Public YOIN_HIN_SHOGO       As String * 2       '「品番照合」の要因 2011.02.03



'-----------------------------------'ホストデータ入出庫区分
Public Const IO_KBN_URI$ = "0"      '売上げ
Public Const IO_KBN_NYU$ = "1"      '入庫
Public Const IO_KBN_SYU$ = "2"      '出庫
Public Const IO_KBN_ZAT$ = "3"      '在庫訂正
Public Const IO_KBN_SYU_JITU$ = "4" '出荷実績
Public Const IO_KBN_HENPIN$ = "5"   '良品返品

Public Const IO_KBN_0$ = "売上げ"
Public Const IO_KBN_1$ = "入　庫"
Public Const IO_KBN_2$ = "出　庫"
Public Const IO_KBN_3$ = "在　訂"
Public Const IO_KBN_4$ = "出荷実"
Public Const IO_KBN_5$ = "良品返"
'-----------------------------------'注文区分
Public Const CYU_KBN_HSP$ = "0"      '補充・スポット
Public Const CYU_KBN_TUK$ = "1"      '月切
Public Const CYU_KBN_SPO$ = "2"      'スポット(読替え注区＝０)
Public Const CYU_KBN_HJU$ = "3"      '補充(読替え注区＝０)
Public Const CYU_KBN_TOK$ = "4"      '特売(読替え注区＝０)
Public Const CYU_KBN_BOU$ = "E"      '貿易
Public Const CYU_KBN_KIN$ = "T"      '特売→緊急（ＷＥＬ専用）

Public Const CYU_KBN_0$ = "補ス"
Public Const CYU_KBN_1$ = "月切"
'''Public Const CYU_KBN_2$ = "スポ"      2003.06.03
Public Const CYU_KBN_2$ = "緊急"        '2003.06.03
Public Const CYU_KBN_3$ = "補充"
Public Const CYU_KBN_4$ = "特売"
'Public Const CYU_KBN_4$ = "一斉"       '2005.11.16 滋賀ＤＣは「一斉」を有効にする

Public Const CYU_KBN_E$ = "貿易"
Public Const CYU_KBN_T$ = "例外"        '2004.05.18
'-----------------------------------'要因関係
Public Const SUM_KBN_IN$ = "1"      '入庫
Public Const SUM_KBN_OT$ = "2"      '出庫
Public Const SUM_KBN_ZT$ = "3"      '在訂±
Public Const SUM_KBN_MV$ = "4"      '移動
Public Const SUM_KBN_NON$ = "0"     'なし

Public Const SUM_KBN_I$ = "入庫　"
Public Const SUM_KBN_O$ = "出庫　"
Public Const SUM_KBN_Z$ = "在訂±"
Public Const SUM_KBN_M$ = "移動　"
Public Const SUM_KBN_N$ = "なし　"

Public Const NORMAL_YOIN$ = "0"     '通常要因
Public Const SYSTEM_YOIN$ = "1"     'システム要因

Public Const NORMAL_YOIN_N$ = "通常　　"
Public Const SYSTEM_YOIN_N$ = "システム"

'-----------------------------------'その他　共通定義
Public Const ETS_MTS$ = "ZZZZZ"     'その他向け先
'-----------------------------------'要因設定
'Public Const ALL_YOIN$ = "0"        'スキャナ／画面使用可
'-----------------------------------'全担当者共通コード
Public Const ALL_TANTO_CODE$ = "ZZZZZ"
