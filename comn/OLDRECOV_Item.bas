Attribute VB_Name = "RECOV_ITEM"
Option Explicit
'********************************************************************
'*
'*              （旧）品目マスタ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Global Const RECOV_ITEM_ID = "RECOV_ITEM"

'ページサイズ
Global Const RECOV_ITEM_PG_SIZ% = 1024

'ポジション・ブロック
Global RECOV_ITEM_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type RECOV_ITEMREC_Tag
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
    HIN_NAME(0 To 24)   As Byte     '品名
    ST_SET_DT(0 To 7)   As Byte     '標準倉庫設定日付
    ST_SOKO(0 To 1)     As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)     As Byte     '             列
    ST_REN(0 To 1)      As Byte     '             連
    ST_DAN(0 To 1)      As Byte     '             段
    BEF_SOKO(0 To 1)    As Byte     '前回入庫倉庫 倉庫
    BEF_RETU(0 To 1)    As Byte     '             列
    BEF_REN(0 To 1)     As Byte     '             連
    BEF_DAN(0 To 1)     As Byte     '             段
    LAST_NYU_DT(0 To 7) As Byte     '最終入庫日付
    LAST_SYU_DT(0 To 7) As Byte     '最終出庫日付
    HIN_NAI(0 To 12)    As Byte     '品番（内部）
    BIKOU_SOKO(0 To 1)  As Byte     '備考 ホスト倉庫
    BIKOU_TANA(0 To 7)  As Byte     '備考 ホスト棚番
    SIZAI_CD(0 To 4)    As Byte     '資材コード
    HOJYU_P(0 To 7)     As Byte     '補充点
    AVE_SYUKA(0 To 7)   As Byte     '月平均出荷数
    SAMPLE_QTY(0 To 0)  As Byte     'サンプル数
    LAST_INP_DT(0 To 7) As Byte     '最終入荷日付
'*------------------------------------------ 2001.02.15 追加 ▽
    LOCK_F(0 To 0)      As Byte     '排他フラグ
    WEL_ID(0 To 1)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
'*------------------------------------------ 2001.02.15 追加 △
    LAST_CHK_DT(0 To 7) As Byte     '最終照合日付2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '最終照合時在庫数2001.06.12
    MOTO_JIGYOBU(0 To 0) As Byte    '元事事業部
    BIKOU(0 To 14)      As Byte     '印刷備考
    IRI_QTY(0 To 7)     As Byte     '印刷入り数
    FILLER(0 To 7)     As Byte      'FILLER
End Type
'データ・バッファ
Global RECOV_ITEMREC As RECOV_ITEMREC_Tag

'キー定義

Type KEY0_RECOV_ITEM            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
End Type


'キー・データ
Global K0_RECOV_ITEM As KEY0_RECOV_ITEM

Type RECOV_ITEM_FSpeck
    fs As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Global RECOV_ITEM_Speck As RECOV_ITEM_FSpeck
 

Function RECOV_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              （旧）品目マスタ  ＯＰＥＮ                          *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    RECOV_ITEM_Open = False
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", RECOV_ITEM_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI 読み込みエラー")
        RECOV_ITEM_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, RECOV_ITEM_POS, RECOV_ITEMREC, Len(RECOV_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "品目マスタ")
                RECOV_ITEM_Open = True
                Exit Function
        End Select
    Loop
End Function


