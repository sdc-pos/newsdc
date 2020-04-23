Attribute VB_Name = "OLD_FURIKAE"
Option Explicit
'********************************************************************
'*
'*              品番振替Ｍ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const OLD_FURIKAE_ID$ = "OLD_FURIKAE"

'ページサイズ
Public Const OLD_FURIKAE_PG_SIZ% = 1024

'ポジション・ブロック
Public OLD_FURIKAE_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OLD_FURIKAEREC_Tag
    HIN_MAE(0 To 19)            As Byte     '振替前品番（外部）
    HIN_GO(0 To 19)             As Byte     '振替後品番（外部）
    BIKOU(0 To 39)              As Byte     '備考
    
    FILLER(0 To 31)             As Byte    '
    
    INS_TANTO(0 To 9)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時

    UPD_TANTO(0 To 9)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type

'データ・バッファ
Public OLD_FURIKAEREC           As OLD_FURIKAEREC_Tag

'キー定義
Type KEY0_OLD_FURIKAE           'ＫＥＹ０
    HIN_MAE(0 To 19)                    As Byte     '振替前品番（外部）
    HIN_GO(0 To 19)                     As Byte     '振替後品番（外部）
End Type

Type KEY1_OLD_FURIKAE           'ＫＥＹ１
    HIN_GO(0 To 19)                     As Byte     '振替後品番（外部）
    HIN_MAE(0 To 19)                    As Byte     '振替前品番（外部）
End Type


'キー・データ
Public K0_OLD_FURIKAE               As KEY0_OLD_FURIKAE
Public K1_OLD_FURIKAE               As KEY1_OLD_FURIKAE

Type OLD_FURIKAE_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck

End Type

Private OLD_FURIKAE_Speck               As OLD_FURIKAE_FSpeck
Private Function OLD_FURIKAE_Create() As Integer
'********************************************************************
'*
'*              品番振替Ｍ　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    OLD_FURIKAE_Create = True
                                            '品番振替Ｍフルパス取込み
    sts = GetIni("FILE", OLD_FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [" & OLD_FURIKAE_ID & "]読み込みエラー")
        Exit Function
    End If
     
    FullPath = RTrim(c)
    
    OLD_FURIKAE_Speck.fs.recoleng = Len(OLD_FURIKAEREC)         ' レコード長
    OLD_FURIKAE_Speck.fs.PageSize = OLD_FURIKAE_PG_SIZ          ' ページサイズ
    OLD_FURIKAE_Speck.fs.idexnumb = 2                   ' インデックス数
    OLD_FURIKAE_Speck.fs.fileflag = 0                   ' ファイルフラグ
    OLD_FURIKAE_Speck.fs.reserve = &H0                  ' 予約済み
'-----------------------------------------------
                                                ' キー０
    OLD_FURIKAE_Speck.ks0.keypos = 1                   ' キーポジション
                                                ' キー長
    OLD_FURIKAE_Speck.ks0.keyleng = 20
                                                ' キーフラグ
    OLD_FURIKAE_Speck.ks0.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    OLD_FURIKAE_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    OLD_FURIKAE_Speck.ks0.reserve = &H0                 ' 予約済み


    OLD_FURIKAE_Speck.ks1.keypos = 21               ' キーポジション
                                                ' キー長
    OLD_FURIKAE_Speck.ks1.keyleng = 20
                                                ' キーフラグ
    OLD_FURIKAE_Speck.ks1.keyflag = BtKfExt  '+ BtKfDup
    OLD_FURIKAE_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    OLD_FURIKAE_Speck.ks1.reserve = &H0                 ' 予約済み


'-----------------------------------------------
                                                ' キー１
    OLD_FURIKAE_Speck.ks2.keypos = 21                   ' キーポジション
    OLD_FURIKAE_Speck.ks2.keyleng = 20                   ' キー長
                                                ' キーフラグ
    OLD_FURIKAE_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfSeg
    OLD_FURIKAE_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    OLD_FURIKAE_Speck.ks2.reserve = &H0                 ' 予約済み

    OLD_FURIKAE_Speck.ks3.keypos = 1                   ' キーポジション
    OLD_FURIKAE_Speck.ks3.keyleng = 20                   ' キー長
                                                ' キーフラグ
    OLD_FURIKAE_Speck.ks3.keyflag = BtKfExt + BtKfDup
    OLD_FURIKAE_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    OLD_FURIKAE_Speck.ks3.reserve = &H0                 ' 予約済み


'-----------------------------------------------

    sts = BTRV(BtOpCreate, OLD_FURIKAE_POS, OLD_FURIKAE_Speck, Len(OLD_FURIKAE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "旧　品番振替Ｍ")
        Exit Function
    End If

    OLD_FURIKAE_Create = False

End Function

Public Function OLD_FURIKAE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品番振替Ｍ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_FURIKAE_Open = True
                                            '品番振替Ｍフルパス取込み
    sts = GetIni("FILE", OLD_FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [OLD_FURIKAE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OLD_FURIKAE_Create()        '品番振替Ｍ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OLD_FURIKAE_POS, OLD_FURIKAEREC, Len(OLD_FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "旧　品番振替Ｍ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "旧　品番振替Ｍ")
                Exit Function
        End Select
    Loop
    OLD_FURIKAE_Open = False
End Function


