Attribute VB_Name = "P_SHURIAGE"
Option Explicit

'********************************************************************
'*
'*              資材売上ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ
Public Const P_SHURIAGE_ID$ = "P_SHURIAGE"

'ページサイズ
Private Const P_SHURIAGE_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SHURIAGE_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_SHURIAGE_REC_Tag
    
    URIAGE_NO(0 To 4)       As Byte         'ﾚｺｰﾄﾞ№
    URIAGE_DT(0 To 7)       As Byte         '売上年月日
    KEIJYO_YM(0 To 5)       As Byte         '計上年月
    TORI_KBN(0 To 0)        As Byte         '取引先区分
    TOKUI_CODE(0 To 4)      As Byte         '得意先ｺｰﾄﾞ
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '資材品番
    G_SYUSHI(0 To 2)        As Byte         '収支単位
    G_HANBAI_KBN(0 To 1)    As Byte         '販売区分
    URIAGE_QTY(0 To 11)     As Byte         '売上数量(S9(8)V99)
    TANKA(0 To 10)          As Byte         '単価(9(8)V99)
    KINGAKU(0 To 8)         As Byte         '売上金額(S9(8))
    SEIKU_F(0 To 0)         As Byte         '請求ﾌﾗｸﾞ
        
    ZEI_KIN(0 To 8)         As Byte         '消費税(S9(8))
    
    FILLER(0 To 19)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SHURIAGE_REC       As P_SHURIAGE_REC_Tag

'キー定義
Public Type KEY0_P_SHURIAGE                 'ＫＥＹ０
    URIAGE_NO(0 To 4)       As Byte         'ﾚｺｰﾄﾞ№
End Type
    
Public Type KEY1_P_SHURIAGE                 'ＫＥＹ１
    KEIJYO_YM(0 To 5)       As Byte         '計上年月
    G_SYUSHI(0 To 2)        As Byte         '収支単位
    TOKUI_CODE(0 To 4)      As Byte         '得意先ｺｰﾄﾞ
    URIAGE_DT(0 To 7)       As Byte         '売上年月日
    URIAGE_NO(0 To 4)       As Byte         'ﾚｺｰﾄﾞ№
End Type
    
    
'キー・データ
Public K0_P_SHURIAGE        As KEY0_P_SHURIAGE
Public K1_P_SHURIAGE        As KEY1_P_SHURIAGE

Type P_SHURIAGE_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_SHURIAGE_Speck    As P_SHURIAGE_FSpeck
Private Function P_SHURIAGE_Create() As Integer
'********************************************************************
'*
'*              資材売上ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SHURIAGE_Create = True
                                            '資材売上ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHURIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURIAGE]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SHURIAGE_Speck.fs.recoleng = Len(P_SHURIAGE_REC)  ' レコード長
    P_SHURIAGE_Speck.fs.PageSize = P_SHURIAGE_PG_SIZ    ' ページサイズ
    P_SHURIAGE_Speck.fs.idexnumb = 2                    ' インデックス数
    P_SHURIAGE_Speck.fs.fileflag = 0                    ' ファイルフラグ
    P_SHURIAGE_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_SHURIAGE_Speck.ks0.keypos = 1                     ' キーポジション
    P_SHURIAGE_Speck.ks0.keyleng = 5                    ' キー長
    P_SHURIAGE_Speck.ks0.keyflag = BtKfExt              ' キーフラグ
    P_SHURIAGE_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks0.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SHURIAGE_Speck.ks1.keypos = 14                    ' キーポジション
    P_SHURIAGE_Speck.ks1.keyleng = 6                    ' キー長
    P_SHURIAGE_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' キーフラグ
    P_SHURIAGE_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks1.reserve = &H0                  ' 予約済み
    
    P_SHURIAGE_Speck.ks2.keypos = 48                    ' キーポジション
    P_SHURIAGE_Speck.ks2.keyleng = 3                    ' キー長
    P_SHURIAGE_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfSeg    ' キーフラグ
    P_SHURIAGE_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks2.reserve = &H0                  ' 予約済み
    
    P_SHURIAGE_Speck.ks3.keypos = 21                    ' キーポジション
    P_SHURIAGE_Speck.ks3.keyleng = 5                    ' キー長
    P_SHURIAGE_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    P_SHURIAGE_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks3.reserve = &H0                  ' 予約済み
    
    P_SHURIAGE_Speck.ks4.keypos = 6                    ' キーポジション
    P_SHURIAGE_Speck.ks4.keyleng = 8                    ' キー長
    P_SHURIAGE_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfSeg   ' キーフラグ
    P_SHURIAGE_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks4.reserve = &H0                  ' 予約済み
    
    P_SHURIAGE_Speck.ks5.keypos = 1                     ' キーポジション
    P_SHURIAGE_Speck.ks5.keyleng = 5                    ' キー長
    P_SHURIAGE_Speck.ks5.keyflag = BtKfExt + BtKfChg             ' キーフラグ
    P_SHURIAGE_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    P_SHURIAGE_Speck.ks5.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー１ △
    
    
    sts = BTRV(BtOpCreate, P_SHURIAGE_POS, P_SHURIAGE_Speck, Len(P_SHURIAGE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材売上ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_SHURIAGE_Create = False

End Function

Public Function P_SHURIAGE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材売上ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SHURIAGE_Open = True
                                            '資材売上ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SHURIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SHURIAGE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SHURIAGE_Create()   '資材売上ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材売上ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材売上ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_SHURIAGE_Open = False

End Function

