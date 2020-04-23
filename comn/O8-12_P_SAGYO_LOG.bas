Attribute VB_Name = "O_P_SAGYO_LOG"
Option Explicit
'********************************************************************
'*
'*              作業実績ﾛｸﾞ  ファイル定義
'*
'*          CREATE 2006.01.30
'********************************************************************
'ファイルＩＤ
Public Const O_P_SAGYO_LOG_ID$ = "O_P_SAGYO_LOG"

'ページサイズ
Public Const O_P_SAGYO_LOG_PG_SIZ% = 1024

'ポジション・ブロック
Public O_P_SAGYO_LOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type O_P_SAGYO_LOG_REC_Tag

    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    WEL_ID(0 To 2)                      As Byte     '対象端末№
    JGYOBU(0 To 0)                      As Byte     '事業部区分
    NAIGAI(0 To 0)                      As Byte     '国内外
    MENU_NO(0 To 1)                     As Byte     'メニューグループ№
    RIRK_ID(0 To 1)                     As Byte     '履歴種別
    ID_NO(0 To 7)                       As Byte     'ID-NO
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
    FILLER(0 To 141)                    As Byte


End Type

'データ・バッファ
Public O_P_SAGYO_LOG_REC      As O_P_SAGYO_LOG_REC_Tag

'キー定義

Type KEY0_O_P_SAGYO_LOG           'ＫＥＹ０
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type

Type KEY1_O_P_SAGYO_LOG           'ＫＥＹ１
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type

Type KEY2_O_P_SAGYO_LOG           'ＫＥＹ２
    TANTO_CODE(0 To 4)                  As Byte     '担当者コード
    MENU_NO(0 To 1)                     As Byte     'メニューグループ№
    JITU_DT(0 To 7)                     As Byte     '実績日付
    JITU_TM(0 To 5)                     As Byte     '実績時刻
End Type



'キー・データ
Public K0_O_P_SAGYO_LOG       As KEY0_O_P_SAGYO_LOG
Public K1_O_P_SAGYO_LOG       As KEY1_O_P_SAGYO_LOG
Public K2_O_P_SAGYO_LOG       As KEY2_O_P_SAGYO_LOG

Type O_P_SAGYO_LOG_FSpeck
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

Private O_P_SAGYO_LOG_Speck   As O_P_SAGYO_LOG_FSpeck
Private Function O_P_SAGYO_LOG_Create() As Integer
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

    O_P_SAGYO_LOG_Create = True
                                            '作業実績ﾛｸﾞフルパス取込み
    sts = GetIni("FILE", O_P_SAGYO_LOG_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SAGYO_LOG]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    O_P_SAGYO_LOG_Speck.fs.recoleng = Len(O_P_SAGYO_LOG_REC)    ' レコード長
    O_P_SAGYO_LOG_Speck.fs.PageSize = O_P_SAGYO_LOG_PG_SIZ      ' ページサイズ
    O_P_SAGYO_LOG_Speck.fs.idexnumb = 3                       ' インデックス数
    O_P_SAGYO_LOG_Speck.fs.fileflag = 0                       ' ファイルフラグ
    O_P_SAGYO_LOG_Speck.fs.reserve = &H0                      ' 予約済み
'------------------------------------------------
                                                            ' キー０
    O_P_SAGYO_LOG_Speck.ks0.keypos = 1                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks0.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks0.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks0.reserve = &H0                     ' 予約済み

    O_P_SAGYO_LOG_Speck.ks1.keypos = 9                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks1.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks1.keyflag = BtKfExt + BtKfDup
    
    O_P_SAGYO_LOG_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks1.reserve = &H0                     ' 予約済み
'------------------------------------------------


'------------------------------------------------
                                                            ' キー１
    O_P_SAGYO_LOG_Speck.ks2.keypos = 15                       ' キーポジション
    O_P_SAGYO_LOG_Speck.ks2.keyleng = 5                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks2.reserve = &H0                     ' 予約済み

    O_P_SAGYO_LOG_Speck.ks3.keypos = 1                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks3.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks3.reserve = &H0                     ' 予約済み


    O_P_SAGYO_LOG_Speck.ks4.keypos = 9                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks4.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks4.keyflag = BtKfExt + BtKfDup
    O_P_SAGYO_LOG_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks4.reserve = &H0                     ' 予約済み

'------------------------------------------------

'------------------------------------------------
                                                            ' キー２
    O_P_SAGYO_LOG_Speck.ks5.keypos = 15                       ' キーポジション
    O_P_SAGYO_LOG_Speck.ks5.keyleng = 5                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks5.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks5.reserve = &H0                     ' 予約済み

    O_P_SAGYO_LOG_Speck.ks6.keypos = 25                       ' キーポジション
    O_P_SAGYO_LOG_Speck.ks6.keyleng = 2                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks6.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks6.reserve = &H0                     ' 予約済み
                                                            
    O_P_SAGYO_LOG_Speck.ks7.keypos = 1                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks7.keyleng = 8                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfSeg
    
    O_P_SAGYO_LOG_Speck.ks7.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks7.reserve = &H0                     ' 予約済み

    O_P_SAGYO_LOG_Speck.ks8.keypos = 9                        ' キーポジション
    O_P_SAGYO_LOG_Speck.ks8.keyleng = 6                       ' キー長
                                                            ' キーフラグ
    O_P_SAGYO_LOG_Speck.ks8.keyflag = BtKfExt + BtKfDup
    
    O_P_SAGYO_LOG_Speck.ks8.keytype = Chr(BtKtString)         ' キータイプ
    O_P_SAGYO_LOG_Speck.ks8.reserve = &H0                     ' 予約済み


'------------------------------------------------




    sts = BTRV(BtOpCreate, O_P_SAGYO_LOG_POS, O_P_SAGYO_LOG_Speck, Len(O_P_SAGYO_LOG_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "作業実績ﾛｸﾞ")
        Exit Function
    End If

    O_P_SAGYO_LOG_Create = False

End Function

Public Function O_P_SAGYO_LOG_Open(Mode As Integer) As Integer
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
    
    O_P_SAGYO_LOG_Open = True
                                            '作業実績ﾛｸﾞフルパス取込み
    sts = GetIni("FILE", O_P_SAGYO_LOG_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SAGYO_LOG]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_P_SAGYO_LOG_POS, O_P_SAGYO_LOG_REC, Len(O_P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_P_SAGYO_LOG_Create()        '作業実績ﾛｸﾞ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_P_SAGYO_LOG_POS, O_P_SAGYO_LOG_REC, Len(O_P_SAGYO_LOG_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    O_P_SAGYO_LOG_Open = False

End Function
