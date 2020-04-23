Attribute VB_Name = "PI00015com"
Option Explicit

Private Type Item_Key_tag
    JGYOBU  As String * 1
    NAIGAI  As String * 1
End Type

Public K_Item_Tbl() As Item_Key_tag   '個装資材品目情報
Public G_Item_Tbl() As Item_Key_tag   '外装資材品目情報



Private Type D_Item_Tbl_Tag
    SYUBETSU    As String * 2               '種別
    JGYOBU      As String * 1               '事業部
    NAIGAI      As String * 1               '国内外
    HIN_GAI     As String * 20              '品番
    QTY         As Double                   '員数
    SHIJI_QTY   As Double                   '数量（指示数）
    BIKOU       As String * 40              '備考（入力値）
    ID_NO       As String * 12              'ID_No(出荷予定ID_No)
End Type



Public D_Item_Tbl()     As D_Item_Tbl_Tag   '同梱／構成品目情報


Public Taget_Key        As String * 8       '更新対象の指図票№

Public Doukon_Tbl_No(0 To 19) _
                        As String * 1

Public Doukon_Start     As Integer          '画面開始行№

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '収支／担当者印刷 OFF:印刷なし ON:印刷あり
Public PRI_MAIN_BCR     As Boolean      'ﾒｲﾝﾊﾞｰｺｰﾄﾞ OFF:印刷なし ON:印刷あり

Public PRI_BIKOU_BCR    As Integer      '備考欄　0：入力値　1:出荷BCR 2:品名

Public PRI_DOUKON       As Boolean      '商品化検査　同梱 OFF:印刷なし ON:印刷あり

Public PRI_NYUKO_IN     As Boolean      '入庫完了印　同梱 OFF:印刷なし ON:印刷あり
Public PRI_INPUT_IN     As Boolean      '入力完了印　同梱 OFF:印刷なし ON:印刷あり

Public PRI_SAGYO_DAY    As Boolean      '作業日／数量／担当 OFF:印刷なし ON:印刷あり 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '下部　品番／№／数量 OFF:印刷なし ON:印刷あり 2007.05.22


Public JISEKI_TITLE     As Variant      '自責の名称タイトル
Public TASEKI_TITLE     As Variant      '他責の名称タイトル

Public HIN_INV          As Boolean      '未登録品番可否


'--------------------------------------------------- 大阪  部材対応　2012.03.20
Public Jyogai_Soko_umu _
                        As Boolean              '除外倉庫設定有無

'--------------------------------------------------- 大阪  部材対応　2012.03.20


'---------------------------------------------- *商品化指図ﾃﾞｰﾀ（親）別ポインタ
'ポジショニング
Public wP_SSHIJI_O_POS  As POSBLK
'データ・バッファ
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'キー・データ
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O

