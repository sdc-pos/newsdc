Attribute VB_Name = "P_SUKEIRE"
Option Explicit

'********************************************************************
'*
'*              商品化指図受入履歴データ  ファイル定義
'*
'*          CREATE 2005.12.14
'********************************************************************
'ファイルＩＤ
Public Const P_SUKEIRE_ID$ = "P_SUKEIRE"

'ページサイズ
Private Const P_SUKEIRE_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SUKEIRE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

Private Type GENKA_TBL_Tag          '原価情報のﾃｰﾌﾞﾙ
    NIN(0 To 2)             As Byte         '人数
    TIMES(0 To 5)           As Byte         '時間
End Type




'レコード定義
Public Type P_SUKEIRE_REC_Tag
    
    SHIJI_NO(0 To 4)       As Byte         '指図票№  未使用とする 2007.11.28
    SEQNO(0 To 2)           As Byte         '追番
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
    UKEIRE_QTY(0 To 10)     As Byte         '受入数量(9(8)V999)
                                            '原価項目
    GENKA_TBL(0 To 9)       As GENKA_TBL_Tag
    
    JISEKI_NAME(0 To 19)    As Byte         '自責要因名
    JISEKI_NIN(0 To 2)      As Byte         '自責  人
    JISEKI_TIMES(0 To 5)    As Byte         '自責  分
    TASEKI_NAME(0 To 19)    As Byte         '他責要因名
    TASEKI_NIN(0 To 2)      As Byte         '他責  人
    TASEKI_TIMES(0 To 5)    As Byte         '他責  分
    
    LAST_F(0 To 0)          As Byte         '最終受入ﾌﾗｸﾞ 0:継続 1:最終
    TORI_CODE(0 To 4)       As Byte         '取引先
    
    'SHIJI_NO(0 To 7)        As Byte         '指図票№   2007.11.28
    FILLER(0 To 94)         As Byte         'Filler     2007.11.28
    
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SUKEIRE_REC        As P_SUKEIRE_REC_Tag

'キー定義

Type KEY0_P_SUKEIRE                         'ＫＥＹ０
'    SHIJI_NO(0 To 4)        As Byte         '指図票№  2007.11.28
    SHIJI_NO(0 To 7)        As Byte         '指図票№   2007.11.28
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
Type KEY1_P_SUKEIRE                         'ＫＥＹ１
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
End Type
    
Type KEY2_P_SUKEIRE                         'ＫＥＹ２
    TORI_CODE(0 To 4)       As Byte         '取引先
    UKEIRE_DT(0 To 7)       As Byte         '受入日
End Type
    
'キー・データ
Public K0_P_SUKEIRE         As KEY0_P_SUKEIRE
Public K1_P_SUKEIRE         As KEY1_P_SUKEIRE
Public K2_P_SUKEIRE         As KEY2_P_SUKEIRE

Type P_SUKEIRE_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SUKEIRE_Speck    As P_SUKEIRE_FSpeck
Private Function P_SUKEIRE_Create() As Integer
'********************************************************************
'*
'*              商品化指図受入履歴ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SUKEIRE_Create = True
                                            '商品化指図受入履歴ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SUKEIRE]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SUKEIRE_Speck.fs.recoleng = Len(P_SUKEIRE_REC)    ' レコード長
    P_SUKEIRE_Speck.fs.PageSize = P_SUKEIRE_PG_SIZ      ' ページサイズ
    P_SUKEIRE_Speck.fs.idexnumb = 3                     ' インデックス数
    P_SUKEIRE_Speck.fs.fileflag = 0                     ' ファイルフラグ
    P_SUKEIRE_Speck.fs.reserve = &H0                    ' 予約済み
    '--------------------------------------------------- キー０ ▽
'2007.11.28    P_SUKEIRE_Speck.ks0.keypos = 1                      ' キーポジション
'2007.11.28    P_SUKEIRE_Speck.ks0.keyleng = 5                     ' キー長
    
    P_SUKEIRE_Speck.ks0.keypos = 184                    ' キーポジション    2007.11.28
    P_SUKEIRE_Speck.ks0.keyleng = 8                     ' キー長            2007.11.28
    
    
    P_SUKEIRE_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    P_SUKEIRE_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks0.reserve = &H0                   ' 予約済み
    
    P_SUKEIRE_Speck.ks1.keypos = 6                      ' キーポジション
    P_SUKEIRE_Speck.ks1.keyleng = 3                     ' キー長
    P_SUKEIRE_Speck.ks1.keyflag = BtKfExt               ' キーフラグ
    P_SUKEIRE_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks1.reserve = &H0                   ' 予約済み
    
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SUKEIRE_Speck.ks2.keypos = 9                      ' キーポジション
    P_SUKEIRE_Speck.ks2.keyleng = 2                     ' キー長
                                                        ' キーフラグ
    P_SUKEIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SUKEIRE_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks2.reserve = &H0                   ' 予約済み
    
    
    P_SUKEIRE_Speck.ks3.keypos = 11                     ' キーポジション
    P_SUKEIRE_Speck.ks3.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    P_SUKEIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SUKEIRE_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks3.reserve = &H0                   ' 予約済み
    
    
    '--------------------------------------------------- キー１ △
    
    '--------------------------------------------------- キー２ ▽
    P_SUKEIRE_Speck.ks4.keypos = 179                    ' キーポジション
    P_SUKEIRE_Speck.ks4.keyleng = 5                     ' キー長
                                                        ' キーフラグ
    P_SUKEIRE_Speck.ks4.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SUKEIRE_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks4.reserve = &H0                   ' 予約済み
    
    
    P_SUKEIRE_Speck.ks5.keypos = 11                     ' キーポジション
    P_SUKEIRE_Speck.ks5.keyleng = 8                     ' キー長
                                                        ' キーフラグ
    P_SUKEIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SUKEIRE_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    P_SUKEIRE_Speck.ks5.reserve = &H0                   ' 予約済み
    
    
    '--------------------------------------------------- キー０ △
    
    
    sts = BTRV(BtOpCreate, P_SUKEIRE_POS, P_SUKEIRE_Speck, Len(P_SUKEIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "商品化指図受入履歴ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SUKEIRE_Create = False

End Function

Public Function P_SUKEIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図受入履歴ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SUKEIRE_Open = True
                                            '商品化指図受入履歴ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SUKEIRE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SUKEIRE_Create()    '商品化指図受入履歴ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "商品化指図受入履歴ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図受入履歴ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SUKEIRE_Open = False

End Function

