Attribute VB_Name = "J_NYU"
Option Explicit
'********************************************************************
'*
'*              入荷チェックデータ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const J_NYU_ID$ = "J_NYU"

'ページサイズ
Public Const J_NYU_PG_SIZ% = 512

'ポジション・ブロック
Public J_NYU_POS    As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type J_NYUREC_Tag
    JGYOBU(0 To 0)      As Byte         '事業部区分
    NAIGAI(0 To 0)      As Byte         '国内外
    HIN_GAI(0 To 19)    As Byte         '品番（外部）
    JITU_QTY(0 To 7)    As Byte         '実績数量
    INS_DATE(0 To 7)    As Byte         '登録日
    FILLER(0 To 25)     As Byte         'FILLER
End Type

'データ・バッファ
Public J_NYUREC         As J_NYUREC_Tag

'キー定義
Type KEY0_J_NYU            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte         '事業部区分
    NAIGAI(0 To 0)      As Byte         '国内外
    HIN_GAI(0 To 19)    As Byte         '品番（外部）
End Type

'キー・データ
Public K0_J_NYU         As KEY0_J_NYU

Type J_NYU_FSpeck
    fs              As BtFileSpeck      'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks1             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
    ks2             As BtKeySpeck       'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private J_NYU_Speck     As J_NYU_FSpeck

Private Function J_NYU_Create() As Integer
'********************************************************************
'*                                                                  *
'*              入荷チェックデータ　ＣＲＥＡＴＥ                    *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    J_NYU_Create = True
                                            '入荷チェックデータフルパス取込み
    sts = GetIni("FILE", J_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [J_NYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    J_NYU_Speck.fs.recoleng = Len(J_NYUREC)     ' レコード長
    J_NYU_Speck.fs.PageSize = J_NYU_PG_SIZ      ' ページサイズ
    J_NYU_Speck.fs.idexnumb = 1                 ' インデックス数
    J_NYU_Speck.fs.fileflag = 0                 ' ファイルフラグ
    J_NYU_Speck.fs.reserve = &H0                ' 予約済み
'------------------------------------------------
                                                ' キー０
    J_NYU_Speck.ks0.keypos = 1                  ' キーポジション
    J_NYU_Speck.ks0.keyleng = 1                 ' キー長
    J_NYU_Speck.ks0.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    J_NYU_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    J_NYU_Speck.ks0.reserve = &H0               ' 予約済み
                                                ' キー０
    J_NYU_Speck.ks1.keypos = 2                  ' キーポジション
    J_NYU_Speck.ks1.keyleng = 1                 ' キー長
    J_NYU_Speck.ks1.keyflag = BtKfExt + BtKfSeg ' キーフラグ
    J_NYU_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    J_NYU_Speck.ks1.reserve = &H0               ' 予約済み
                                                ' キー０
    J_NYU_Speck.ks2.keypos = 3                  ' キーポジション
    J_NYU_Speck.ks2.keyleng = 20                ' キー長
    J_NYU_Speck.ks2.keyflag = BtKfExt           ' キーフラグ
    J_NYU_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    J_NYU_Speck.ks2.reserve = &H0               ' 予約済み
'------------------------------------------------

    sts = BTRV(BtOpCreate, J_NYU_POS, J_NYU_Speck, Len(J_NYU_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入荷チェックデータ")
        Exit Function
    End If
    
    J_NYU_Create = False

End Function
Public Function J_NYU_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              入荷チェックデータ　ＯＰＥＮ                        *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    J_NYU_Open = True
                                        '入荷チェックデータフルパス取込み
    sts = GetIni("FILE", J_NYU_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [J_NYU]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, J_NYU_POS, J_NYUREC, Len(J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = J_NYU_Create()        '入荷チェックデータ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, J_NYU_POS, J_NYUREC, Len(J_NYUREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入荷チェックデータ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "入荷チェックデータ")
                Exit Function
        End Select
    Loop

    J_NYU_Open = False

End Function


