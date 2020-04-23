Attribute VB_Name = "OLD_IDO2"
Option Explicit
'********************************************************************
'*
'*              在庫移動歴　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_IDO2_ID$ = "OLD_IDO2"

'ページサイズ
Public Const OLD_IDO2_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_IDO2_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_IDO2REC_Tag
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
    JGYOBU(0 To 0)                      As Byte     '事業部区分
    NAIGAI(0 To 0)                      As Byte     '国内外
    HIN_GAI(0 To 19)                    As Byte     '品番（外部）
    RIRK_ID(0 To 1)                     As Byte     '履歴種別
    SUMI_JITU_QTY(0 To 7)               As Byte     '実績数量(商品化済み)
    MI_JITU_QTY(0 To 7)                 As Byte     '実績数量(未商品)
    FROM_SOKO(0 To 1)                   As Byte     'From 倉庫№
    FROM_RETU(0 To 1)                   As Byte     '   　列
    FROM_REN(0 To 1)                    As Byte     '   　連
    FROM_DAN(0 To 1)                    As Byte     '   　段
    TO_SOKO(0 To 1)                     As Byte     'ＴＯ 倉庫№
    TO_RETU(0 To 1)                     As Byte     '   　列
    TO_REN(0 To 1)                      As Byte     '   　連
    TO_DAN(0 To 1)                      As Byte     '   　段
    DEN_DT(0 To 7)                      As Byte     '伝票日付
    DEN_NO(0 To 9)                      As Byte     '伝票№
    PRG_ID(0 To 7)                      As Byte     '出力元プログラム
    HIN_NAI(0 To 19)                    As Byte     '品番（内部）
    NYUKA_DT(0 To 7)                    As Byte     '入荷日付
    NYUKO_DT(0 To 7)                    As Byte     '入庫日付
    WEL_ID(0 To 2)                      As Byte     '対象端末№
    RIRK_NAME(0 To 9)                   As Byte     '履歴種別名称
    HIN_NAME(0 To 24)                   As Byte     '品名
    SUMI_HIN_Zaiko_Qty(0 To 7)          As Byte     '品目別在庫数（商品化済み）
    MI_HIN_Zaiko_Qty(0 To 7)            As Byte     '品目別在庫数（未商品）
    SUMI_FROM_TANA_Zaiko_Qty(0 To 7)    As Byte     'FROM棚別品目別在庫数
    SUMI_TO_TANA_Zaiko_Qty(0 To 7)      As Byte     'TO棚別品目別在庫数
    MI_FROM_TANA_Zaiko_Qty(0 To 7)      As Byte     'FROM棚別品目別在庫数
    MI_TO_TANA_Zaiko_Qty(0 To 7)        As Byte     'TO棚別品目別在庫数
    TOKU_MARK(0 To 0)                   As Byte     '特売りマーク
    MEMO(0 To 59)                       As Byte     'メモ
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    TANTO_NAME(0 To 19)                 As Byte     '担当者名称
    MUKE_CODE(0 To 7)                   As Byte     '得意先コード
    MUKE_NAME(0 To 39)                  As Byte     '得意先名称
    SS_CODE(0 To 7)                     As Byte     '直送先コード
    SS_NAME(0 To 39)                    As Byte     '直送先名称
    MUKE_DNAME(0 To 9)                  As Byte     '得意先略称
    MUKE_CHG_CD(0 To 1)                 As Byte     '向け先読替えコード
    SUM_KBN(0 To 0)                     As Byte     '集計区分
    ID_NO(0 To 7)                       As Byte     'ID-NO
    FILLER(0 To 90)                     As Byte
    
End Type

'データ・バッファ
Public OLD_IDO2REC   As OLD_IDO2REC_Tag

'キー定義
Type KEY0_OLD_IDO2            'ＫＥＹ０
    JGYOBU(0 To 0)              As Byte     '事業部区分
    JITU_DT(0 To 7)             As Byte     '実績日付
    JITU_TM(0 To 5)             As Byte     '実績時刻
End Type

'キー・データ
Public K0_OLD_IDO2                   As KEY0_OLD_IDO2

Public Function OLD_IDO2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              在庫移動歴　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_IDO2_Open = True
                                            '在庫移動歴フルパス取込み
    sts = GetIni("FILE", OLD_IDO2_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [OLD_IDO2]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫移動歴")
                Exit Function
        End Select
    Loop
    OLD_IDO2_Open = False
End Function


