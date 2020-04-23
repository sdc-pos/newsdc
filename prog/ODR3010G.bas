Attribute VB_Name = "ODR3010G"
Option Explicit
'********************************************************************
'*
'*              ＯＤＲ３０１０用　共通変数
'*
'********************************************************************
Public ODR30102_Return As Integer         '確認画面終了状態
Public ODR30104_Return As Integer         '確認画面終了状態
Public ODR30105_Return As Integer         '確認画面終了状態


Public KIBOU_DT         As String       '希望納期


Public DIS_ITEM         As String       '子部品コード
Public DIS_ITEM_NM      As String       '子部品名
Public DIS_USE_QTY      As String       '使用数量
Public DIS_MRP_QTY      As String       '必要数
Public DIS_ZAI_QTY      As String       '月初在庫
Public DIS_FUSOKU       As String       '不足数
Public DIS_ORDR_QTY     As String       '注文数
Public DIS_ZAN_QTY      As String       '仕入残

Public DIS_HANSEIHIN_QTY _
                        As String       '半製品数
Public DIS_TEI_QTY      As String       '在訂±


Public DIS_LOT_QTY      As String       'ロット数
Public DIS_SECT_CD      As String       '仕入先
Public DIS_SECT_NM      As String       '仕入先名
Public DIS_TANKA        As String       '仕入単価
Public DIS_KIBOU_DT     As String       '希望納期
Public DIS_KAITO_DE     As String       '回答納期
Public DIS_KEY          As String       'ＫＥＹ項目


Public DIS_DELI_CD      As String       '納入先コード
Public DIS_DELI_NM      As String       '納入先名

Public DIS_Item_Zaiko      As String    '前月末在庫
Public DIS_ZAIKO_ODR      As String     '在庫＋発注数
Public DIS_ZAIKO_UKE      As String     '在庫＋仕入済数

Public Key_SIMUKE       As String       '仕向け先
Public Key_JIGYOBU      As String       '事業部
Public Key_NAIGAI       As String       '国内外
Public Key_USE_YM       As String       '使用月（YYYYMM)
Public Key_INS_NO       As String       '登録順
Public Key_HinGai      As String        '親品番
Public Key_ORDER_NO     As String       '親品番　注文№
Public Key_BUN_NO       As String       '分納回数

Public pubBikou_1   As String           '備考１
Public pubBikou_2   As String           '備考２
Public pubBikou_3   As String           '備考３


'グリッド用定義
Public ORDR_GRID   As New XArrayDB

'Public Const Col_No% = 0                '行№

Public Const Col_ITEM% = 0              '子部品コード
Public Const Col_ITEM_NM% = 1           '子部品名
Public Const Col_USE_QTY% = 2           '使用数量
Public Const Col_MRP_QTY% = 3           '必要数
Public Const Col_ZAI_QTY% = 4           '月初在庫
Public Const Col_FUSOKU% = 5            '不足数
Public Const Col_ORDR_QTY% = 6          '注文数

Public Const Col_ZAN_QTY% = 7           '仕入残

Public Const Col_HANSEIHIN_QTY% = 8     '半製品数

Public Const Col_TEI_QTY% = 9          '在訂±


Public Const Col_LOT_QTY% = 10          'ロット数
Public Const Col_KAITO_DT% = 11         '回答納期
Public Const Col_KIBOU_DT% = 12         '希望納期

Public Const Col_SECT_CD% = 13           '仕入先
Public Const Col_SECT_NM% = 14          '仕入先名
Public Const Col_TANKA% = 15            '仕入単価
Public Const Col_KEY% = 16              '使用月
Public Const Col_JIGYOBU% = 17          '事業部
Public Const Col_NAIGAI% = 18           '国内外

Public Const Col_DELI_CD% = 19          '納入先
Public Const Col_DELI_NM% = 20          '納入先名

Public Const Col_Item_Zaiko% = 21       '品目Ｍ在庫数
Public Const Col_ZAIKO_ODR% = 22        '在庫＋親＜０数
Public Const Col_ZAIKO_UKE% = 23        '在庫＋仕入済数

'ステーション№
Public WS_NO       As String * 10
'---------------------------------------------- *検索用資材注文ﾃﾞｰﾀ
'ポジショニング
Public wP_SHORDER_POS       As POSBLK
'データ・バッファ
Public wP_SHORDER_REC       As P_SHORDER_REC_Tag
'キー・データ
Public K2_wP_SHORDER        As KEY2_P_SHORDER



Public Function wP_SHORDER_Open(Mode As Integer) As Integer
'****************************************************
'*      「資材注文ﾃﾞｰﾀ」    ＯＰＥＮ処理
'*
'*  資材注文ﾃﾞｰﾀを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wP_SHORDER_Open = True
                                    '資材注文ﾃﾞｰﾀ　フルパス取込み
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHORDER]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wP_SHORDER_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    wP_SHORDER_Open = False

End Function

Public Function wP_SHORDER_CLOSE() As Integer

'****************************************************
'*      「資材注文ﾃﾞｰﾀ」    ＣＬＯＳＥ処理
'*
'*  資材注文ﾃﾞｰﾀを別ポインタでＣＬＯＳＥする
'*  (呼び元で終了時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'****************************************************
Dim sts As Integer
    
    wP_SHORDER_CLOSE = True
    
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
            Exit Function
    End Select

    wP_SHORDER_CLOSE = False

End Function

