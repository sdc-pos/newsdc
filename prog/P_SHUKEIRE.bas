Attribute VB_Name = "P_SHUKEIRE"
Option Explicit

'********************************************************************
'*
'*              資材受入履歴ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SHUKEIRE_ID$ = "P_SHUKEIRE"

'ページサイズ
Private Const P_SHUKEIRE_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SHUKEIRE_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_SHUKEIRE_REC_Tag
    
    ORDER_NO(0 To 4)        As Byte         '注文№
    SEQNO(0 To 2)           As Byte         '追番
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)       As Byte         '受入日
    UKEIRE_QTY(0 To 11)     As Byte         '受入数量(S9(8)V99)
    UKEIRE_TANKA(0 To 10)   As Byte         '受入単価(9(8)V99)
    UKEIRE_KINGAKU(0 To 8)  As Byte         '受入金額(S9(8))
    LAST_F(0 To 0)          As Byte         '最終受入ﾌﾗｸﾞ 0:継続 1:最終
    KEIJYO_YM(0 To 5)       As Byte         '計上年月(YYYYMM)
    ZEI_KIN(0 To 8)         As Byte         '消費税額(S9(8))    2007.04.29
    FILLER(0 To 44)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SHUKEIRE_REC       As P_SHUKEIRE_REC_Tag

'キー定義

Public Type KEY0_P_SHUKEIRE                     'ＫＥＹ０
    ORDER_NO(0 To 4)        As Byte         '注文№
    SEQNO(0 To 2)           As Byte         '追番
End Type

Public Type KEY1_P_SHUKEIRE                     'ＫＥＹ１
    KEIJYO_YM(0 To 5)       As Byte         '計上年月(YYYYMM)
    ORDER_CODE(0 To 4)      As Byte         '注文先ｺｰﾄﾞ
    UKEIRE_DT(0 To 7)        As Byte         '注文日
End Type
    
Public Type KEY2_P_SHUKEIRE                     'ＫＥＹ２
    KEIJYO_YM(0 To 5)       As Byte         '計上年月(YYYYMM)
    UKEIRE_DT(0 To 7)        As Byte         '注文日
End Type
    
    
'キー・データ
Public K0_P_SHUKEIRE        As KEY0_P_SHUKEIRE
Public K1_P_SHUKEIRE        As KEY1_P_SHUKEIRE
Public K2_P_SHUKEIRE        As KEY2_P_SHUKEIRE

Type P_SHUKEIRE_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SHUKEIRE_Speck    As P_SHUKEIRE_FSpeck
Private Function P_SHUKEIRE_Create() As Integer
'********************************************************************
'*
'*              資材受入履歴ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHUKEIRE_Create = True
                                            '資材受入履歴ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHUKEIRE]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHUKEIRE_Speck.fs.recoleng = Len(P_SHUKEIRE_REC)  ' レコード長
    P_SHUKEIRE_Speck.fs.PageSize = P_SHUKEIRE_PG_SIZ    ' ページサイズ
    P_SHUKEIRE_Speck.fs.idexnumb = 3                    ' インデックス数
    P_SHUKEIRE_Speck.fs.fileflag = 0                    ' ファイルフラグ
    P_SHUKEIRE_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHUKEIRE_Speck.ks0.keypos = 1                     ' キーポジション
    P_SHUKEIRE_Speck.ks0.keyleng = 5                    ' キー長
    P_SHUKEIRE_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    P_SHUKEIRE_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks0.reserve = &H0                  ' 予約済み
    
    P_SHUKEIRE_Speck.ks1.keypos = 6                     ' キーポジション
    P_SHUKEIRE_Speck.ks1.keyleng = 3                    ' キー長
    P_SHUKEIRE_Speck.ks1.keyflag = BtKfExt              ' キーフラグ
    P_SHUKEIRE_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks1.reserve = &H0                  ' 予約済み
    
    
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SHUKEIRE_Speck.ks2.keypos = 55                    ' キーポジション
    P_SHUKEIRE_Speck.ks2.keyleng = 6                    ' キー長
                                                        ' キーフラグ
    P_SHUKEIRE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks2.reserve = &H0                  ' 予約済み
    
    P_SHUKEIRE_Speck.ks3.keypos = 9                     ' キーポジション
    P_SHUKEIRE_Speck.ks3.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    P_SHUKEIRE_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks3.reserve = &H0                  ' 予約済み
    
    P_SHUKEIRE_Speck.ks4.keypos = 14                    ' キーポジション
    P_SHUKEIRE_Speck.ks4.keyleng = 8                    ' キー長
    P_SHUKEIRE_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup   ' キーフラグ
    P_SHUKEIRE_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks4.reserve = &H0                  ' 予約済み
    
    
    '--------------------------------------------------- キー２ △
    
    
    '--------------------------------------------------- キー１ ▽
    P_SHUKEIRE_Speck.ks5.keypos = 55                    ' キーポジション
    P_SHUKEIRE_Speck.ks5.keyleng = 6                    ' キー長
                                                        ' キーフラグ
    P_SHUKEIRE_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SHUKEIRE_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks5.reserve = &H0                  ' 予約済み
    
    
    P_SHUKEIRE_Speck.ks6.keypos = 14                    ' キーポジション
    P_SHUKEIRE_Speck.ks6.keyleng = 8                    ' キー長
    P_SHUKEIRE_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup   ' キーフラグ
    P_SHUKEIRE_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    P_SHUKEIRE_Speck.ks6.reserve = &H0                  ' 予約済み
    
    
    '--------------------------------------------------- キー２ △
    
    
    sts = BTRV(BtOpCreate, P_SHUKEIRE_POS, P_SHUKEIRE_Speck, Len(P_SHUKEIRE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材受入ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SHUKEIRE_Create = False

End Function

Public Function P_SHUKEIRE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材受入ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHUKEIRE_Open = True
                                            '資材受入ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHUKEIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHUKEIRE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHUKEIRE_Create()   '資材受入ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材受入ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材受入ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SHUKEIRE_Open = False

End Function

