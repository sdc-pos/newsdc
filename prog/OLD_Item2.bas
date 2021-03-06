Attribute VB_Name = "OLD_ITEM2"
Option Explicit
'********************************************************************
'*
'*              品目マスタ  ファイル定義
'*
'*          CREATE 2004.02.19
'********************************************************************
'ファイルＩＤ
Public Const OLD_ITEM2_ID$ = "OLD_ITEM2"

'ページサイズ
Public Const OLD_ITEM2_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_ITEM2_POS         As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_ITEM2REC_Tag
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
    WEL_ID(0 To 2)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
'*------------------------------------------ 2001.02.15 追加 △
    LAST_CHK_DT(0 To 7) As Byte     '最終照合日付2001.06.12
    LAST_CHK_QTY(0 To 7) As Byte    '最終照合時在庫数2001.06.12
    MOTO_JIGYOBU(0 To 0) As Byte    '元事事業部     '未使用2004.02
    BIKOU(0 To 14)      As Byte     '印刷備考
    IRI_QTY(0 To 7)     As Byte     '印刷入り数
    
    JAN_CODE(0 To 12)   As Byte     'Janコード      2004.02
    HIN_CHANGE(0 To 12) As Byte     '品番読み替え   2004.02
    GOODS_KBN(0 To 0)   As Byte     '商品化有無     2004.02
    PACKING_NO(0 To 3)  As Byte     '個装箱��       2004.02
    
    FILLER(0 To 167)    As Byte     'FILLER         2004.02
End Type
'データ・バッファ
Public OLD_ITEM2REC As OLD_ITEM2REC_Tag

'キー定義

Type KEY0_OLD_ITEM2            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
End Type



'キー・データ
Public K0_OLD_ITEM2 As KEY0_OLD_ITEM2

Public Function OLD_ITEM2_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_ITEM2_Open = True
                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", OLD_ITEM2_ID, "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [OLD_ITEM2]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "品目マスタ")
                Exit Function
        End Select
    Loop

    OLD_ITEM2_Open = False

End Function


