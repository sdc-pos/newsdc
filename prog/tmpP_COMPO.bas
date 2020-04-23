Attribute VB_Name = "tmpP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              構成マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const tmpP_COMPO_ID$ = "tmpP_COMPO"

'ページサイズ
Private Const tmpP_COMPO_PG_SIZ% = 512

'ポジション・ブロック
Public tmpP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type tmpP_COMPOREC_Tag
    
    
    SHIMUKE(0 To 2)         As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    KO_SYUBETSU(0 To 1)     As Byte         '子　種別
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    KO_QTY(0 To 5)          As Byte         '子　員数(999V99)
    KO_BIKOU(0 To 39)       As Byte         '子　備考
    FILLER(0 To 137)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public tmpP_COMPOREC         As tmpP_COMPOREC_Tag

'キー定義

Type KEY0_tmpP_COMPO                        'ＫＥＹ０
    SHIMUKE(0 To 2)         As Byte         '仕向け先
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '親品番
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
'キー・データ
Public K0_tmpP_COMPO        As KEY0_tmpP_COMPO

Type tmpP_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private tmpP_COMPO_Speck    As tmpP_COMPO_FSpeck

Private Function tmpP_COMPO_Create() As Integer
'********************************************************************
'*                                                                  *
'*              構成マスタ(一時ファイル)ＣＲＥＡＴＥ                *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    tmpP_COMPO_Create = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", tmpP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpP_COMPO]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    tmpP_COMPO_Speck.fs.recoleng = Len(tmpP_COMPOREC)       ' レコード長
    tmpP_COMPO_Speck.fs.PageSize = tmpP_COMPO_PG_SIZ        ' ページサイズ
    tmpP_COMPO_Speck.fs.idexnumb = 1                        ' インデックス数
    tmpP_COMPO_Speck.fs.fileflag = 0                        ' ファイルフラグ
    tmpP_COMPO_Speck.fs.reserve = &H0                       ' 予約済み
    '--------------------------------------------------- キー０ ▽
    tmpP_COMPO_Speck.ks0.keypos = 1                         ' キーポジション
    tmpP_COMPO_Speck.ks0.keyleng = 3                        ' キー長
    tmpP_COMPO_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpP_COMPO_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks0.reserve = &H0                      ' 予約済み
    
    tmpP_COMPO_Speck.ks1.keypos = 4                         ' キーポジション
    tmpP_COMPO_Speck.ks1.keyleng = 1                        ' キー長
    tmpP_COMPO_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpP_COMPO_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks1.reserve = &H0                      ' 予約済み
    
    tmpP_COMPO_Speck.ks2.keypos = 5                         ' キーポジション
    tmpP_COMPO_Speck.ks2.keyleng = 1                        ' キー長
    tmpP_COMPO_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpP_COMPO_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks2.reserve = &H0                      ' 予約済み
    
    tmpP_COMPO_Speck.ks3.keypos = 6                         ' キーポジション
    tmpP_COMPO_Speck.ks3.keyleng = 20                       ' キー長
    tmpP_COMPO_Speck.ks3.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpP_COMPO_Speck.ks3.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks3.reserve = &H0                      ' 予約済み
    
    tmpP_COMPO_Speck.ks4.keypos = 26                        ' キーポジション
    tmpP_COMPO_Speck.ks4.keyleng = 1                        ' キー長
    tmpP_COMPO_Speck.ks4.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    tmpP_COMPO_Speck.ks4.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks4.reserve = &H0                      ' 予約済み
    
    tmpP_COMPO_Speck.ks5.keypos = 27                        ' キーポジション
    tmpP_COMPO_Speck.ks5.keyleng = 3                        ' キー長
    tmpP_COMPO_Speck.ks5.keyflag = BtKfExt                  ' キーフラグ
    tmpP_COMPO_Speck.ks5.keytype = Chr(BtKtString)          ' キータイプ
    tmpP_COMPO_Speck.ks5.reserve = &H0                      ' 予約済み
    
    '--------------------------------------------------- キー０ △
    sts = BTRV(BtOpCreate, tmpP_COMPO_POS, tmpP_COMPO_Speck, Len(tmpP_COMPO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "構成マスタ（一時ファイル）")
        Exit Function
    End If
    
    tmpP_COMPO_Create = False

End Function

Public Function tmpP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              構成マスタ（一時ファイル）  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim ans         As Integer


    tmpP_COMPO_Open = True
                                            '構成マスタフルパス取込み
    sts = GetIni("FILE", tmpP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [tmpP_COMPO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, tmpP_COMPO_POS, tmpP_COMPOREC, Len(tmpP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                
                ans = MsgBox("他端末で、作業用ファイル使用中です。", vbRetryCancel, "確認入力")
                
                If ans = vbCancel Then
                    Exit Function
                End If
                Sleep (500&)
            Case BtErrFileNotFound
                sts = tmpP_COMPO_Create()      '構成マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, tmpP_COMPO_POS, tmpP_COMPOREC, Len(tmpP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "構成マスタ（一時ファイル）")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "構成マスタ（一時ファイル）")
                Exit Function
        End Select
    Loop
    
    tmpP_COMPO_Open = False

End Function

