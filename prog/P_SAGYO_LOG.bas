Attribute VB_Name = "P_SAGYO_LOG"
Option Explicit
'********************************************************************
'*
'*              作業実績ﾛｸﾞ  ファイル定義
'*
'*          CREATE 2006.01.30
'********************************************************************
'ファイルＩＤ
Public Const P_SAGYO_LOG_ID$ = "P_SAGYO_LOG"

'ページサイズ
Public Const P_SAGYO_LOG_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SAGYO_LOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type P_SAGYO_LOG_REC_Tag

    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    WEL_ID(0 To 2)                      As Byte     '対象端末№
    JGYOBU(0 To 0)                      As Byte     '事業部区分
    NAIGAI(0 To 0)                      As Byte     '国内外
    MENU_NO(0 To 1)                     As Byte     'メニューグループ№
    RIRK_ID(0 To 1)                     As Byte     '履歴種別
'    ID_NO(0 To 7)                       As Byte     'ID-NO
    ID_NO(0 To 11)                      As Byte     'ID-NO (8桁→12桁)      2006/05/24
    HIN_GAI(0 To 19)                    As Byte     '品番（外部）
    SUMI_JITU_QTY(0 To 7)               As Byte     '実績数量(商品化済み)
    MI_JITU_QTY(0 To 7)                 As Byte     '実績数量(未商品)
    MUKE_CODE(0 To 7)                   As Byte     '得意先コード
    SS_CODE(0 To 7)                     As Byte     '直送先コード
    FROM_SOKO(0 To 1)                   As Byte     'From 倉庫№
    FROM_RETU(0 To 1)                   As Byte     '   　列
    FROM_REN(0 To 1)                    As Byte     '   　連
    FROM_DAN(0 To 1)                    As Byte     '   　段
    TO_SOKO(0 To 1)                     As Byte     'ＴＯ 倉庫№
    TO_RETU(0 To 1)                     As Byte     '   　列
    TO_REN(0 To 1)                      As Byte     '   　連
    TO_DAN(0 To 1)                      As Byte     '   　段
    PRG_ID(0 To 9)                      As Byte     '出力元プログラム
    WORK_TM(0 To 5)                     As Byte     '作業時間(秒)・・・追加：2008.08.19
    
    
    SHIJI_No(0 To 7)                    As Byte     '指図票№   未使用とする 2010.09.03
    
    
    HIN_CHECK_LABEL_CNT(0 To 2)         As Byte     '品番ﾁｪｯｸﾗﾍﾞﾙ件数   2010.09.03
    HIN_CHECK_GENPIN_CNT(0 To 2)        As Byte     '品番ﾁｪｯｸ現品票件数   2010.09.03
    
    JAN_CODE(0 To 19)                   As Byte     'JANｺｰﾄﾞ            2011.08.18
    
    MEMO(0 To 19)                       As Byte     'メモ               2014.07.01
    
    
    HIN_CHECK_GAISOU_CNT(0 To 2)        As Byte     '品番ﾁｪｯｸ外装件数   2015.11.07
    
    
    
    
    FILLER(0 To 74)                     As Byte     '                   2011.08.18 117-->97 2014.07.01 97-->77 2015.11.07 77-->74






End Type

'データ・バッファ
Public P_SAGYO_LOG_REC      As P_SAGYO_LOG_REC_Tag

'キー定義

Type KEY0_P_SAGYO_LOG           'ＫＥＹ０
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type

Type KEY1_P_SAGYO_LOG           'ＫＥＹ１
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type

Type KEY2_P_SAGYO_LOG           'ＫＥＹ２
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    MENU_NO(0 To 1)                     As Byte     'メニューグループ№
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type



'キー・データ
Public K0_P_SAGYO_LOG       As KEY0_P_SAGYO_LOG
Public K1_P_SAGYO_LOG       As KEY1_P_SAGYO_LOG
Public K2_P_SAGYO_LOG       As KEY2_P_SAGYO_LOG

Type P_SAGYO_LOG_FSpeck
    fs  As BtFileSpeck                      'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck                       'ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
    ks3 As BtKeySpeck
    ks4 As BtKeySpeck
    ks5 As BtKeySpeck
    ks6 As BtKeySpeck
    ks7 As BtKeySpeck
    ks8 As BtKeySpeck
End Type

Private P_SAGYO_LOG_Speck   As P_SAGYO_LOG_FSpeck
Private Function P_SAGYO_LOG_Create() As Integer
'********************************************************************
'*
'*              作業実績ﾛｸﾞ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2006.01.30
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    P_SAGYO_LOG_Create = True
                                            '作業実績ﾛｸﾞフルパス取込み
    sts = GetIni("FILE", P_SAGYO_LOG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SAGYO_LOG]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    P_SAGYO_LOG_Speck.fs.recoleng = Len(P_SAGYO_LOG_REC)    ' レコード長
    P_SAGYO_LOG_Speck.fs.PageSize = P_SAGYO_LOG_PG_SIZ      ' ページサイズ
    P_SAGYO_LOG_Speck.fs.idexnumb = 3                       ' インデックス数
    P_SAGYO_LOG_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_SAGYO_LOG_Speck.fs.reserve = &H0                      ' 予約済み
'------------------------------------------------
                                                            ' キー０
    P_SAGYO_LOG_Speck.ks0.keypos = 1                        ' キーポジション
    P_SAGYO_LOG_Speck.ks0.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks0.reserve = &H0                     ' 予約済み

    P_SAGYO_LOG_Speck.ks1.keypos = 9                        ' キーポジション
    P_SAGYO_LOG_Speck.ks1.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks1.keyflag = BtKfExt + BtKfDup
    
    P_SAGYO_LOG_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks1.reserve = &H0                     ' 予約済み
'------------------------------------------------


'------------------------------------------------
                                                            ' キー１
    P_SAGYO_LOG_Speck.ks2.keypos = 15                       ' キーポジション
    P_SAGYO_LOG_Speck.ks2.keyleng = 5                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks2.reserve = &H0                     ' 予約済み

    P_SAGYO_LOG_Speck.ks3.keypos = 1                        ' キーポジション
    P_SAGYO_LOG_Speck.ks3.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks3.reserve = &H0                     ' 予約済み


    P_SAGYO_LOG_Speck.ks4.keypos = 9                        ' キーポジション
    P_SAGYO_LOG_Speck.ks4.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks4.keyflag = BtKfExt + BtKfDup
    P_SAGYO_LOG_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks4.reserve = &H0                     ' 予約済み

'------------------------------------------------

'------------------------------------------------
                                                            ' キー２
    P_SAGYO_LOG_Speck.ks5.keypos = 15                       ' キーポジション
    P_SAGYO_LOG_Speck.ks5.keyleng = 5                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks5.reserve = &H0                     ' 予約済み

    P_SAGYO_LOG_Speck.ks6.keypos = 25                       ' キーポジション
    P_SAGYO_LOG_Speck.ks6.keyleng = 2                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks6.reserve = &H0                     ' 予約済み
                                                            
    P_SAGYO_LOG_Speck.ks7.keypos = 1                        ' キーポジション
    P_SAGYO_LOG_Speck.ks7.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    P_SAGYO_LOG_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks7.reserve = &H0                     ' 予約済み

    P_SAGYO_LOG_Speck.ks8.keypos = 9                        ' キーポジション
    P_SAGYO_LOG_Speck.ks8.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    P_SAGYO_LOG_Speck.ks8.keyflag = BtKfExt + BtKfDup
    
    P_SAGYO_LOG_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    P_SAGYO_LOG_Speck.ks8.reserve = &H0                     ' 予約済み


'------------------------------------------------




    sts = BTRV(BtOpCreate, P_SAGYO_LOG_POS, P_SAGYO_LOG_Speck, Len(P_SAGYO_LOG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "作業実績ﾛｸﾞ")
        Exit Function
    End If

    P_SAGYO_LOG_Create = False

End Function

Public Function P_SAGYO_LOG_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              作業実績ﾛｸﾞ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    P_SAGYO_LOG_Open = True
                                            '作業実績ﾛｸﾞフルパス取込み
    sts = GetIni("FILE", P_SAGYO_LOG_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SAGYO_LOG]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SAGYO_LOG_Create()        '作業実績ﾛｸﾞ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "作業実績ﾛｸﾞ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "作業実績ﾛｸﾞ")
                Exit Function
        End Select
    Loop
    
    P_SAGYO_LOG_Open = False

End Function
