Attribute VB_Name = "OLD_STOCK"
Option Explicit
'********************************************************************
'*
'*              （旧）棚卸しデータ  ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_STOCK_ID$ = "OLD_STOCK"

'ページサイズ
Public Const OLD_STOCK_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_STOCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_STOCKREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 12)        As Byte     '品番（外部）
    ST_LOCATION(0 To 7)     As Byte     '標準入庫倉庫
    HOST_ZAIKO(0 To 7)      As Byte     '松下理論在庫
    POS_ZAIKO(0 To 7)       As Byte     'ＰＯＳ総在庫
    ST_ZAIKO(0 To 7)        As Byte     '標準棚番在庫
    
    EE1_LOCATION(0 To 7)    As Byte     '別置き１
    EE1_ZAIKO(0 To 7)       As Byte     '在庫
    EE2_LOCATION(0 To 7)    As Byte     '別置き２
    EE2_ZAIKO(0 To 7)       As Byte     '在庫
    EE3_LOCATION(0 To 7)    As Byte     '別置き３
    EE3_ZAIKO(0 To 7)       As Byte     '在庫
    
    ETC_ZAIKO(0 To 7)       As Byte     'その他在庫
    CHECK_MARK(0 To 0)      As Byte     '照合マーク
    PRINT_YMD(0 To 7)       As Byte     '印刷日付
    INPUT_YMD(0 To 7)       As Byte     '入力日付
    
    SAI_QTY(0 To 8)         As Byte     '差異数　2004.06.29
    
    FILLER(0 To 30)         As Byte
    
End Type
'データ・バッファ
Public OLD_STOCKREC         As OLD_STOCKREC_Tag

'キー定義

Type KEY0_OLD_STOCK                     'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 12)        As Byte     '品番（外部）
End Type


'キー・データ
Public K0_OLD_STOCK         As KEY0_OLD_STOCK
Function OLD_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              棚卸しデータ  ＯＰＥＮ                              *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_STOCK_Open = True
                                    '棚卸しデータフルパス取込み
    sts = GetIni("FILE", OLD_STOCK_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI [OLD_STOCK]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_STOCK_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "(旧)棚卸しデータ")
                Exit Function
        End Select
    Loop
    
    OLD_STOCK_Open = False

End Function
