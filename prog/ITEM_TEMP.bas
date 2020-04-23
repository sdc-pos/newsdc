Attribute VB_Name = "ITEM_TEMP"
Option Explicit
'********************************************************************
'*                                                                  *
'*              品目一時ファイル  ファイル定義                      *
'*                                                                  *
'*          CREATE 2008.02.03                                       *
'********************************************************************
'ファイルＩＤ
Public Const ITEM_TEMP_ID$ = "ITEM_TEMP"

'ページサイズ
Public Const ITEM_TEMP_PG_SIZ% = 512

'ポジション・ブロック
Public ITEM_TEMP_POS        As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type ITEM_TEMP_REC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）
    
    KO_HIN_GAI(0 To 16)     As Byte         '個装資材
    CLASS(0 To 3)           As Byte         'ｸﾗｽ

    ST_SOKO(0 To 1)         As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte         '             列
    ST_REN(0 To 1)          As Byte         '             連
    ST_DAN(0 To 1)          As Byte         '             段
    

End Type

'データ・バッファ
Public ITEM_TEMP_REC        As ITEM_TEMP_REC_Tag

'キー定義

Type KEY0_ITEM_TEMP             'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）
End Type

Type KEY1_ITEM_TEMP             'ＫＥＹ１
    KO_HIN_GAI(0 To 16)     As Byte         '個装資材
    
    ST_SOKO(0 To 1)         As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte         '             列
    ST_REN(0 To 1)          As Byte         '             連
    ST_DAN(0 To 1)          As Byte         '             段
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）

End Type

Type KEY2_ITEM_TEMP             'ＫＥＹ２

    CLASS(0 To 3)           As Byte         'ｸﾗｽ
    
    ST_SOKO(0 To 1)         As Byte         '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte         '             列
    ST_REN(0 To 1)          As Byte         '             連
    ST_DAN(0 To 1)          As Byte         '             段
    
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番（外部）

End Type


'キー・データ
Public K0_ITEM_TEMP         As KEY0_ITEM_TEMP
Public K1_ITEM_TEMP         As KEY1_ITEM_TEMP
Public K2_ITEM_TEMP         As KEY2_ITEM_TEMP

Type ITEM_TEMP_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public ITEM_TEMP_Speck      As ITEM_TEMP_FSpeck
 
Private Function ITEM_TEMP_Create() As Integer
'********************************************************************
'*                                                                  *
'*              品目一時ファイル  ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2008.02.3                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ITEM_TEMP_Create = True
                                            '担当者マスタフルパス取込み
    sts = GetIni("FILE", ITEM_TEMP_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    ITEM_TEMP_Speck.fs.recoleng = Len(ITEM_TEMP_REC)    ' レコード長
    ITEM_TEMP_Speck.fs.PageSize = ITEM_TEMP_PG_SIZ%     ' ページサイズ
    ITEM_TEMP_Speck.fs.idexnumb = 3                     ' インデックス数
    ITEM_TEMP_Speck.fs.fileflag = 0                     ' ファイルフラグ
    ITEM_TEMP_Speck.fs.reserve = &H0                    ' 予約済み
                                                        
    
    '---------------------------------------------------' キー０
    ITEM_TEMP_Speck.ks0.keypos = 1                          ' キーポジション
    ITEM_TEMP_Speck.ks0.keyleng = 22                        ' キー長
    ITEM_TEMP_Speck.ks0.keyflag = BtKfExt                   ' キーフラグ
    ITEM_TEMP_Speck.ks0.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks0.reserve = &H0                       ' 予約済み
    '---------------------------------------------------' キー０

    '---------------------------------------------------' キー１
    ITEM_TEMP_Speck.ks1.keypos = 23                         ' キーポジション
    ITEM_TEMP_Speck.ks1.keyleng = 20                        ' キー長
    ITEM_TEMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    ITEM_TEMP_Speck.ks1.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks1.reserve = &H0                       ' 予約済み
    
    ITEM_TEMP_Speck.ks2.keypos = 47                         ' キーポジション
    ITEM_TEMP_Speck.ks2.keyleng = 8                         ' キー長
    ITEM_TEMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    ITEM_TEMP_Speck.ks2.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks2.reserve = &H0                       ' 予約済み
    
    ITEM_TEMP_Speck.ks3.keypos = 47                         ' キーポジション
    ITEM_TEMP_Speck.ks3.keyleng = 8                         ' キー長
    ITEM_TEMP_Speck.ks3.keyflag = BtKfExt                   ' キーフラグ
    ITEM_TEMP_Speck.ks3.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks3.reserve = &H0                       ' 予約済み
    '---------------------------------------------------' キー１

    '---------------------------------------------------' キー２
    ITEM_TEMP_Speck.ks1.keypos = 43                         ' キーポジション
    ITEM_TEMP_Speck.ks1.keyleng = 4                         ' キー長
    ITEM_TEMP_Speck.ks1.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    ITEM_TEMP_Speck.ks1.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks1.reserve = &H0                       ' 予約済み
    
    ITEM_TEMP_Speck.ks2.keypos = 47                         ' キーポジション
    ITEM_TEMP_Speck.ks2.keyleng = 8                         ' キー長
    ITEM_TEMP_Speck.ks2.keyflag = BtKfExt + BtKfSeg         ' キーフラグ
    ITEM_TEMP_Speck.ks2.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks2.reserve = &H0                       ' 予約済み
    
    ITEM_TEMP_Speck.ks3.keypos = 1                          ' キーポジション
    ITEM_TEMP_Speck.ks3.keyleng = 20                        ' キー長
    ITEM_TEMP_Speck.ks3.keyflag = BtKfExt                   ' キーフラグ
    ITEM_TEMP_Speck.ks3.keytype = Chr(BtKtString)           ' キータイプ
    ITEM_TEMP_Speck.ks3.reserve = &H0                       ' 予約済み
    '---------------------------------------------------' キー２





    sts = BTRV(BtOpCreate, ITEM_TEMP_POS, ITEM_TEMP_Speck, Len(ITEM_TEMP_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品目一時データ")
    End If

    ITEM_TEMP_Create = False

End Function

Function ITEM_TEMP_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              品目一時データ  ＯＰＥＮ                          　*
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    ITEM_TEMP_Open = True
                                            '品目一時データフルパス取込み
    sts = GetIni("FILE", ITEM_TEMP_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, ITEM_TEMP_POS, ITEM_TEMP_REC, Len(ITEM_TEMP_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_TEMP_Create()    '品目一時データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_TEMP_POS, ITEM_TEMP_REC, Len(ITEM_TEMP_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品目一時データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品目一時データ")
                Exit Function
        End Select
    Loop

    ITEM_TEMP_Open = False
    
End Function
