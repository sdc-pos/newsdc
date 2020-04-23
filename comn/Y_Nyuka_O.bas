Attribute VB_Name = "Y_NYU_O"
Option Explicit
'********************************************************************
'*                                                                  *
'*              入荷予定データ（大阪PC向け）  ファイル定義          *
'*                                                                  *
'********************************************************************
'ファイルＩＤ
Public Const Y_NYU_O_ID$ = "Y_NYU_O"

'ページサイズ
Public Const Y_NYU_O_PG_SIZ% = 2048

'ポジション・ブロック
Public Y_NYU_O_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type Y_NYUREC_O_Tag
    JGYOBU(0 To 0)              As Byte     '事業部
    SOKO_NO(0 To 1)             As Byte     '倉庫№
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
    NYUKO_YMD(0 To 7)           As Byte     '入庫日(入荷日)
    DEN_NO(0 To 5)              As Byte     '伝票№
    MAKER_CODE(0 To 5)          As Byte     'ﾒｰｶｰｺｰﾄﾞ
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品番
    Y_SURYO(0 To 7)             As Byte     '予定数量
    J_SURYO(0 To 7)             As Byte     '実績数量
    TANTO_CODE(0 To 4)          As Byte     '担当者ｺｰﾄﾞ
    ORDER_NO(0 To 9)            As Byte     '注文№
    KENPIN_F(0 To 0)            As Byte     '検品F
    WEL_ID(0 To 2)              As Byte     '使用子機ID
    PRG_ID(0 To 7)              As Byte     '使用中プログラム
    FILLER(0 To 165)            As Byte     'FILLER
    
End Type

'データ・バッファ
Public Y_NYU_O_REC                  As Y_NYUREC_O_Tag

'キー定義
Type KEY0_Y_NYU_O            'ＫＥＹ０
    SEQ_NO(0 To 2)              As Byte     'SEQ_NO
End Type

Type KEY1_Y_NYU_O            'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_NO(0 To 19)             As Byte     '品番
End Type

Type KEY2_Y_NYU_O            'ＫＥＹ１
    WEL_ID(0 To 2)              As Byte     '使用子機ID
    PRG_ID(0 To 7)              As Byte     '使用中プログラム
End Type



'キー・データ
Public K0_Y_NYU_O               As KEY0_Y_NYU_O
Public K1_Y_NYU_O               As KEY1_Y_NYU_O
Public K2_Y_NYU_O               As KEY2_Y_NYU_O

Private Type Y_NYU_O_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
End Type

Private Y_NYU_O_Speck As Y_NYU_O_FSpeck

Private Function Y_NYU_O_Create() As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ(大阪PC向け)  ＣＲＥＡＴＥ            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Y_NYU_O_Create = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", Y_NYU_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_NYU_O]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    Y_NYU_O_Speck.fs.recoleng = Len(Y_NYU_O_REC)    ' レコード長
    Y_NYU_O_Speck.fs.PageSize = Y_NYU_O_PG_SIZ      ' ページサイズ
    Y_NYU_O_Speck.fs.idexnumb = 3                   ' インデックス数
    Y_NYU_O_Speck.fs.fileflag = 0                   ' ファイルフラグ
    Y_NYU_O_Speck.fs.reserve = &H0                  ' 予約済み
    '-------------------------------------------
                                                
    Y_NYU_O_Speck.ks0.keypos = 4                    ' キーポジション
    Y_NYU_O_Speck.ks0.keyleng = 3                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks0.keyflag = BtKfExt
    Y_NYU_O_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks0.reserve = &H0                 ' 予約済み
                                                
                                                
                                                ' キー０
    '-------------------------------------------
    
    '-------------------------------------------
                                                ' キー１
    Y_NYU_O_Speck.ks1.keypos = 1                    ' キーポジション
    Y_NYU_O_Speck.ks1.keyleng = 1                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks1.reserve = &H0                 ' 予約済み
                                                
    Y_NYU_O_Speck.ks2.keypos = 27                    ' キーポジション
    Y_NYU_O_Speck.ks2.keyleng = 1                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks2.reserve = &H0                 ' 予約済み
                                                
    Y_NYU_O_Speck.ks3.keypos = 28                    ' キーポジション
    Y_NYU_O_Speck.ks3.keyleng = 20                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_O_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks3.reserve = &H0                 ' 予約済み
                                                
                                                
                                                ' キー１
    '-------------------------------------------
    
    
    '-------------------------------------------
                                                ' キー２
    Y_NYU_O_Speck.ks4.keypos = 80                   ' キーポジション
    Y_NYU_O_Speck.ks4.keyleng = 3                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    Y_NYU_O_Speck.ks4.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks4.reserve = &H0                 ' 予約済み
                                                
    Y_NYU_O_Speck.ks5.keypos = 83                   ' キーポジション
    Y_NYU_O_Speck.ks5.keyleng = 8                   ' キー長
                                                    ' キーフラグ
    Y_NYU_O_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup
    Y_NYU_O_Speck.ks5.keytype = Chr(BtKtString)     ' キータイプ
    Y_NYU_O_Speck.ks5.reserve = &H0                 ' 予約済み
                                                
                                                
                                                ' キー２
    '-------------------------------------------
    
    
    sts = BTRV(BtOpCreate, Y_NYU_O_POS, Y_NYU_O_Speck, Len(Y_NYU_O_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入荷予定データ")
        Y_NYU_O_Create = True
        Exit Function
    End If

    Y_NYU_O_Create = False

End Function

Function Y_NYU_O_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              入荷予定データ(大阪PC向け)  ＯＰＥＮ                *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    Y_NYU_O_Open = True
                                            '入荷予定データフルパス取込み
    sts = GetIni("FILE", Y_NYU_O_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [Y_NYU]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Y_NYU_O_Create()        '入荷予定データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入荷予定データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "入荷予定データ")
                Exit Function
        End Select
    Loop
    
    Y_NYU_O_Open = False

End Function


