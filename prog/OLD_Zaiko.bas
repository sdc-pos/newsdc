Attribute VB_Name = "OLD_ZAIKO"
Option Explicit
'********************************************************************
'*
'*              （旧）在庫データ ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_ZAIKO_ID$ = "OLD_ZAIKO"

'ページサイズ
Public Const OLD_ZAIKO_PG_SIZ% = 2048

'ポジション・ブロック
Public OLD_ZAIKO_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
Type OLD_ZAIKOREC_Tag
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
    NYUKO_DT(0 To 7)    As Byte     '入庫日付
    HIN_NAI(0 To 12)    As Byte     '品番（内部）
    YUKO_Z_QTY(0 To 7)  As Byte     '有効在庫数
    LOCK_F(0 To 0)      As Byte     '排他フラグ
    WEL_ID(0 To 2)      As Byte     '使用子機ID
    PRG_ID(0 To 7)      As Byte     '使用中プログラム
    GOODS_YMD(0 To 7)   As Byte     '商品化日付
    FILLER(0 To 46)     As Byte     'FILLER
End Type

'データ・バッファ
Public OLD_ZAIKOREC     As OLD_ZAIKOREC_Tag

'キー定義
Type KEY0_OLD_ZAIKO                 'ＫＥＹ０
    Soko_No(0 To 1)     As Byte     '倉庫№
    Retu(0 To 1)        As Byte     '棚番　列
    Ren(0 To 1)         As Byte     '棚番　連
    Dan(0 To 1)         As Byte     '棚番　段
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 12)    As Byte     '品番（外部）
    GOODS_ON(0 To 0)    As Byte     '商品化／未商品化
    NYUKA_DT(0 To 7)    As Byte     '入荷日付
End Type


'キー・データ
Public K0_OLD_ZAIKO         As KEY0_OLD_ZAIKO
Public Function OLD_ZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              （旧）在庫データ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_ZAIKO_Open = True
                                            '在庫データ　フルパス取込み
    sts = GetIni("FILE", OLD_ZAIKO_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_ZAIKO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_ZAIKO_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(旧)在庫データ")
                Exit Function
        End Select
    Loop
    OLD_ZAIKO_Open = False

End Function

