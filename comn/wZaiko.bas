Attribute VB_Name = "wZAIKO"
Option Explicit
'********************************************************************
'*
'*              在庫データ(ﾜｰｸ) ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Global Const wZAIKO_ID = "wZAIKO"

'ページサイズ
Global Const wZAIKO_PG_SIZ% = 1024

'ポジション・ブロック
Global wZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type wZAIKOREC_Tag
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    '2005.12.05 13-->20
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    NYUKO_DT(0 To 7)    As Byte     '入庫日付
    '2005.12.05 13-->20
    HIN_NAI(0 To 19)    As Byte     '品番（内部）
    YUKO_Z_QTY(0 To 7)  As Byte     '有効在庫数
    LOCK_F(0 To 0)      As Byte     '排他フラグ
    WEL_ID(0 To 2)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
    GOODS_YMD(0 To 7)   As Byte     '商品化日付
    
    '2005.12.05 項目追加
    SHIIRE_CODE(0 To 4) As Byte     '仕入先ｺｰﾄﾞ
    SHIIRE_TANKA(0 To 10) As Byte   '仕入単価(9(8)V99)
    KEIJYO_YM(0 To 5)   As Byte     '計上年月
    '2005.12.05 項目追加
    
    FILLER(0 To 74)     As Byte     'FILLER
End Type

'データ・バッファ
Global wZAIKOREC As wZAIKOREC_Tag

'キー定義
Type KEY0_wZAIKO                    'ＫＥＹ０
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
End Type

Type KEY1_wZAIKO                    'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
End Type

Type KEY2_wZAIKO                    'ＫＥＹ２
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
End Type

Type KEY3_wZAIKO                    'ＫＥＹ３
    WEL_ID(0 To 1)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
End Type

Type KEY4_wZAIKO                     'ＫＥＹ４
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
End Type

'キー・データ
Global K0_wZAIKO As KEY0_wZAIKO
Global K1_wZAIKO As KEY1_wZAIKO
Global K2_wZAIKO As KEY2_wZAIKO
Global K3_wZAIKO As KEY3_wZAIKO
Global K4_wZAIKO As KEY4_wZAIKO

Function wZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              在庫データ　ＯＰＥＮ                                *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    wZAIKO_Open = False
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", wZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        wZAIKO_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫データ(ﾜｰｸ)")
                wZAIKO_Open = True
                Exit Function
        End Select
    Loop
End Function

