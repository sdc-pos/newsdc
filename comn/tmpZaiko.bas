Attribute VB_Name = "tmpZAIKO"
Option Explicit
'********************************************************************
'*
'*              在庫データ（一時データ） ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const tmpZAIKO_ID$ = "tmpZAIKO"

'ページサイズ
Public Const tmpZAIKO_PG_SIZ% = 1024

'ポジション・ブロック
Public tmpZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type tmpZAIKOREC_Tag
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
Public tmpZAIKOREC      As tmpZAIKOREC_Tag

'キー定義

Type KEY0_tmpZAIKO                    'ＫＥＹ０
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
End Type

Type KEY1_tmpZAIKO                     'ＫＥＹ１
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
End Type

'キー・データ
Public K0_tmpZAIKO      As KEY0_tmpZAIKO
Public K1_tmpZAIKO      As KEY1_tmpZAIKO

Type tmpZAIKO_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
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
    ks15    As BtKeySpeck
End Type

Private tmpZAIKO_Speck As tmpZAIKO_FSpeck
Public Function tmpZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              在庫データ（一時データ）　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    tmpZAIKO_Open = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", tmpZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpZAIKO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫データ（一時データ）")
                Exit Function
        End Select
    Loop
    tmpZAIKO_Open = False

End Function

