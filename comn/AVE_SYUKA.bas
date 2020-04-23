Attribute VB_Name = "AVE_SYUKA"
Option Explicit
'********************************************************************
'*
'*              平均出荷数　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const AVE_SYUKA_ID$ = "AVE_SYUKA"

'ページサイズ
Public Const AVE_SYUKA_PG_SIZ% = 512

'ポジション・ブロック
Public AVE_SYUKA_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type AVE_SYUKAREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品目コード（外部）
    ST_LOCATION(0 To 7)     As Byte         '標準棚番
    UPDATE_YMD(0 To 7)      As Byte         '集計年月日
    ZEN3_YM(0 To 5)         As Byte         '前々々年月         2011.07.01 未使用 ｽﾍﾟｰｽｸﾘｱｰ
    ZEN3_SYUKA(0 To 7)      As Byte         '前々々月出荷数     2011.07.01 未使用 0ｸﾘｱｰ
    ZEN2_YM(0 To 5)         As Byte         '前々年月           2011.07.01 未使用 ｽﾍﾟｰｽｸﾘｱｰ
    ZEN2_SYUKA(0 To 7)      As Byte         '前々月出荷数       2011.07.01 未使用 0ｸﾘｱｰ
    ZEN1_YM(0 To 5)         As Byte         '前年月             2011.07.01 未使用 ｽﾍﾟｰｽｸﾘｱｰ
    ZEN1_SYUKA(0 To 7)      As Byte         '前月出荷数         2011.07.01 全出荷として使用
    AVE_SYUKA(0 To 7)       As Byte         '平均出荷数
    Two_Year_SYUKA(0 To 7)  As Byte         '過去２年間実績

'-------------------------------------------' 2011.07.01 ▼
    TOTAL_CNT(0 To 7)           As Byte     '総出荷件数
    TOTAL_AVE_CNT(0 To 7)       As Byte     '平均総出荷件数


    S_SYUKA_QTY1(0 To 7)        As Byte     '生産計画出荷数(1)
    S_SYUKA_CNT1(0 To 7)        As Byte     '生産計画出荷件数(1)
    S_AVE_SYUKA_QTY1(0 To 7)    As Byte     '平均生産計画出荷数(1)
    S_AVE_SYUKA_CNT1(0 To 7)    As Byte     '平均生産計画出荷件数(1)

    S_SYUKA_QTY2(0 To 7)        As Byte     '生産計画出荷数(2)
    S_SYUKA_CNT2(0 To 7)        As Byte     '生産計画出荷件数(2)
    S_AVE_SYUKA_QTY2(0 To 7)    As Byte     '平均生産計画出荷数(2)
    S_AVE_SYUKA_CNT2(0 To 7)    As Byte     '平均生産計画出荷件数(2)


    NAI_BUHIN(0 To 0)           As Byte     '国内供給部品区分
    HIN_NAME(0 To 39)           As Byte     '品名


    FILLER(0 To 38)             As Byte     'FILLER
'-------------------------------------------' 2011.07.01　▲





End Type

'データ・バッファ
Public AVE_SYUKAREC         As AVE_SYUKAREC_Tag

'キー定義
Type KEY0_AVE_SYUKA         'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品目コード（外部）
End Type

Type KEY1_AVE_SYUKA         'ＫＥＹ０
    ST_LOCATION(0 To 7)     As Byte         '標準棚番
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品目コード（外部）
End Type

'キー・データ
Public K0_AVE_SYUKA         As KEY0_AVE_SYUKA
Public K1_AVE_SYUKA         As KEY1_AVE_SYUKA

Type AVE_SYUKA_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks2     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks3     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks4     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks5     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks6     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
    ks7     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private AVE_SYUKA_Speck As AVE_SYUKA_FSpeck

Private Function AVE_SYUKA_Create() As Integer
'********************************************************************
'*
'*              月平均出荷数　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    AVE_SYUKA_Create = True
                                            '月平均出荷数フルパス取込み
    sts = GetIni("FILE", AVE_SYUKA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [AVE_SYUKA]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    AVE_SYUKA_Speck.fs.recoleng = Len(AVE_SYUKAREC)     ' レコード長
    AVE_SYUKA_Speck.fs.PageSize = AVE_SYUKA_PG_SIZ      ' ページサイズ
    AVE_SYUKA_Speck.fs.idexnumb = 2                     ' インデックス数
    AVE_SYUKA_Speck.fs.fileflag = 0                     ' ファイルフラグ
    AVE_SYUKA_Speck.fs.reserve = &H0                    ' 予約済み
                                                    
'---------------------------------------------------
                                                        ' キー０
    AVE_SYUKA_Speck.ks0.keypos = 1                      ' キーポジション
    AVE_SYUKA_Speck.ks0.keyleng = 1                     ' キー長
    AVE_SYUKA_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    AVE_SYUKA_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks0.reserve = &H0                   ' 予約済み

    AVE_SYUKA_Speck.ks1.keypos = 2                      ' キーポジション
    AVE_SYUKA_Speck.ks1.keyleng = 1                     ' キー長
    AVE_SYUKA_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    AVE_SYUKA_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks1.reserve = &H0                   ' 予約済み

    AVE_SYUKA_Speck.ks2.keypos = 3                      ' キーポジション
    AVE_SYUKA_Speck.ks2.keyleng = 20                    ' キー長
    AVE_SYUKA_Speck.ks2.keyflag = BtKfExt               ' キーフラグ
    AVE_SYUKA_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks2.reserve = &H0                   ' 予約済み
'---------------------------------------------------
                                                        ' キー１
    AVE_SYUKA_Speck.ks3.keypos = 23                     ' キーポジション
    AVE_SYUKA_Speck.ks3.keyleng = 8                     ' キー長
    AVE_SYUKA_Speck.ks3.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    AVE_SYUKA_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks3.reserve = &H0                   ' 予約済み

    AVE_SYUKA_Speck.ks4.keypos = 1                      ' キーポジション
    AVE_SYUKA_Speck.ks4.keyleng = 1                     ' キー長
    AVE_SYUKA_Speck.ks4.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    AVE_SYUKA_Speck.ks4.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks4.reserve = &H0                   ' 予約済み

    AVE_SYUKA_Speck.ks5.keypos = 2                      ' キーポジション
    AVE_SYUKA_Speck.ks5.keyleng = 1                     ' キー長
    AVE_SYUKA_Speck.ks5.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    AVE_SYUKA_Speck.ks5.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks5.reserve = &H0                   ' 予約済み

    AVE_SYUKA_Speck.ks6.keypos = 3                      ' キーポジション
    AVE_SYUKA_Speck.ks6.keyleng = 20                    ' キー長
    AVE_SYUKA_Speck.ks6.keyflag = BtKfExt               ' キーフラグ
    AVE_SYUKA_Speck.ks6.keytype = Chr(BtKtString)       ' キータイプ
    AVE_SYUKA_Speck.ks6.reserve = &H0                   ' 予約済み

    sts = BTRV(BtOpCreate, AVE_SYUKA_POS, AVE_SYUKA_Speck, Len(AVE_SYUKA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "月平均出荷数")
        Exit Function
    End If

    AVE_SYUKA_Create = False

End Function

Public Function AVE_SYUKA_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              月平均出荷数　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    AVE_SYUKA_Open = True
                                            '月平均出荷数フルパス取込み
    sts = GetIni("FILE", AVE_SYUKA_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [AVE_SYUKA]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = AVE_SYUKA_Create()    '月平均出荷数作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "月平均出荷数")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "月平均出荷数")
                Exit Function
        End Select
    Loop

    AVE_SYUKA_Open = False

End Function
