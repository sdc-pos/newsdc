Attribute VB_Name = "ABC"
Option Explicit
'********************************************************************
'*                                                                  *
'*              ＡＢＣ管理集計ファイル（一時ファイル） ファイル定義 *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
'ファイルＩＤ
Public Const ABC_ID = "ABC"

'ページサイズ
Public Const ABC_PG_SIZ% = 512

'ポジション・ブロック
Public ABC_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type ABCREC_Tag
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    ST_LOCATION(0 To 7) As Byte     '標準棚番
    PACKING_NO(0 To 3)  As Byte     '個装箱№
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    RANK_NOW(0 To 2)    As Byte     '現在設定ランク
    RANK_NEW(0 To 2)    As Byte     '新ランク

End Type
'データ・バッファ
Public ABCREC           As ABCREC_Tag


'キー定義
Type KEY0_ABC                       'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    ST_LOCATION(0 To 7) As Byte     '標準棚番
    PACKING_NO(0 To 3)  As Byte     '個装箱№
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type
    
'キー・データ
Public K0_ABC           As KEY0_ABC

Private Type ABC_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private ABC_Speck    As ABC_FSpeck
Private Function ABC_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ABC管理集計ファイル  ＣＲＥＡＴＥ                   *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ABC_Create = True
                                            'ABC管理集計ファイルフルパス取込み
    sts = GetIni("FILE", ABC_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[ABC] 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    ABC_Speck.fs.recoleng = Len(ABCREC)             ' レコード長
    ABC_Speck.fs.PageSize = ABC_PG_SIZ              ' ページサイズ
    ABC_Speck.fs.idexnumb = 1                       ' インデックス数
    ABC_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ABC_Speck.fs.reserve = &H0                      ' 予約済み
                                                    
'---------------------------------------------------' キー０
    ABC_Speck.ks0.keypos = 1                        ' キーポジション
    ABC_Speck.ks0.keyleng = 1                       ' キー長
    ABC_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    ABC_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ABC_Speck.ks0.reserve = &H0                     ' 予約済み

    ABC_Speck.ks1.keypos = 2                        ' キーポジション
    ABC_Speck.ks1.keyleng = 1                       ' キー長
    ABC_Speck.ks1.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    ABC_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ABC_Speck.ks1.reserve = &H0                     ' 予約済み

    ABC_Speck.ks2.keypos = 3                        ' キーポジション
    ABC_Speck.ks2.keyleng = 8                       ' キー長
    ABC_Speck.ks2.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    ABC_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    ABC_Speck.ks2.reserve = &H0                     ' 予約済み

    ABC_Speck.ks3.keypos = 11                       ' キーポジション
    ABC_Speck.ks3.keyleng = 4                       ' キー長
    ABC_Speck.ks3.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    ABC_Speck.ks3.keytype = Chr(BtKtString)         ' キータイプ
    ABC_Speck.ks3.reserve = &H0                     ' 予約済み

    ABC_Speck.ks4.keypos = 15                       ' キーポジション
    ABC_Speck.ks4.keyleng = 20                      ' キー長
    ABC_Speck.ks4.keyflag = BtKfExt                 ' キーフラグ
    ABC_Speck.ks4.keytype = Chr(BtKtString)         ' キータイプ
    ABC_Speck.ks4.reserve = &H0                     ' 予約済み

    sts = BTRV(BtOpCreate, ABC_POS, ABC_Speck, Len(ABC_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ＡＢＣ管理集計ファイル")
        Exit Function
    End If
    
    ABC_Create = False

End Function

Function ABC_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ＡＢＣ管理集計ファイル  ＯＰＥＮ                    *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.22                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ABC_Open = True
                                            'ＡＢＣ管理集計ファイルフルパス取込み
    sts = GetIni("FILE", ABC_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, ABC_POS, ABCREC, Len(ABCREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ABC_Create()        'ＡＢＣ管理集計ファイル作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ABC_POS, ABCREC, Len(ABCREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ＡＢＣ管理集計ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ＡＢＣ管理集計ファイル")
                Exit Function
        End Select
    Loop
    ABC_Open = False
End Function
