Attribute VB_Name = "O_P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              商品化指図データ（子）  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ

Public Const O_P_SSHIJI_K_ID$ = "O_P_SSHIJI_K"

'ページサイズ
Private Const O_P_SSHIJI_K_PG_SIZ% = 512

'ポジション・ブロック
Public O_P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Public Type O_P_SSHIJI_K_REC_Tag
    
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    KO_SYUBETSU(0 To 1)     As Byte         '子　種別
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    KO_QTY(0 To 5)          As Byte         '子　員数(999V99)
    KO_SHIJI_QTY(0 To 10)   As Byte         '指示数(9(8)V99)
    KO_BIKOU(0 To 39)       As Byte         '子　備考
    KO_ID_NO(0 To 7)        As Byte         '子 ＩＤ＿ＮＯ
    CALCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
    FILLER(0 To 64)         As Byte         'Filler
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public O_P_SSHIJI_K_REC       As O_P_SSHIJI_K_REC_Tag

'キー定義

Type KEY0_O_P_SSHIJI_K                        'ＫＥＹ０
    SHIJI_NO(0 To 4)        As Byte         '指図票№
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
    
Type KEY1_O_P_SSHIJI_K                        'ＫＥＹ０
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_ID_NO(0 To 7)        As Byte         '子 ＩＤ＿ＮＯ
End Type
    
    
'キー・データ
Public K0_O_P_SSHIJI_K        As KEY0_O_P_SSHIJI_K

Type O_P_SSHIJI_K_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private O_P_SSHIJI_K_Speck    As O_P_SSHIJI_K_FSpeck
Private Function O_P_SSHIJI_K_Create() As Integer
'********************************************************************
'*
'*              商品化指図データ（子）  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    O_P_SSHIJI_K_Create = True
                                            '手配指図データ（子）フルパス取込み
    sts = GetIni("FILE", O_P_SSHIJI_K_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SSHIJI_K]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    O_P_SSHIJI_K_Speck.fs.recoleng = Len(O_P_SSHIJI_K_REC)  ' レコード長
    O_P_SSHIJI_K_Speck.fs.PageSize = O_P_SSHIJI_K_PG_SIZ    ' ページサイズ
    O_P_SSHIJI_K_Speck.fs.idexnumb = 2                    ' インデックス数
    O_P_SSHIJI_K_Speck.fs.fileflag = 0                    ' ファイルフラグ
    O_P_SSHIJI_K_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
    O_P_SSHIJI_K_Speck.ks0.keypos = 1                     ' キーポジション
    O_P_SSHIJI_K_Speck.ks0.keyleng = 5                    ' キー長
    O_P_SSHIJI_K_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    O_P_SSHIJI_K_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    O_P_SSHIJI_K_Speck.ks0.reserve = &H0                  ' 予約済み
    
    O_P_SSHIJI_K_Speck.ks1.keypos = 6                     ' キーポジション
    O_P_SSHIJI_K_Speck.ks1.keyleng = 1                    ' キー長
    O_P_SSHIJI_K_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    O_P_SSHIJI_K_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    O_P_SSHIJI_K_Speck.ks1.reserve = &H0                  ' 予約済み
    
    
    O_P_SSHIJI_K_Speck.ks2.keypos = 7                     ' キーポジション
    O_P_SSHIJI_K_Speck.ks2.keyleng = 3                    ' キー長
    O_P_SSHIJI_K_Speck.ks2.keyflag = BtKfExt              ' キーフラグ
    O_P_SSHIJI_K_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    O_P_SSHIJI_K_Speck.ks2.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    O_P_SSHIJI_K_Speck.ks3.keypos = 1                     ' キーポジション
    O_P_SSHIJI_K_Speck.ks3.keyleng = 5                    ' キー長
                                                        ' キーフラグ
    O_P_SSHIJI_K_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    O_P_SSHIJI_K_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    O_P_SSHIJI_K_Speck.ks3.reserve = &H0                  ' 予約済み
    
    O_P_SSHIJI_K_Speck.ks4.keypos = 6                     ' キーポジション
    O_P_SSHIJI_K_Speck.ks4.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    O_P_SSHIJI_K_Speck.ks4.keyflag = BtKfExt + BtKfDup
    O_P_SSHIJI_K_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    O_P_SSHIJI_K_Speck.ks4.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー１ △
    
    
    sts = BTRV(BtOpCreate, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_Speck, Len(O_P_SSHIJI_K_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "手配指図データ（子）")
        Exit Function
    End If
    
    O_P_SSHIJI_K_Create = False

End Function

Public Function O_P_SSHIJI_K_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図データ（子）  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    O_P_SSHIJI_K_Open = True
                                            '手配指図データ（子）フルパス取込み
    sts = GetIni("FILE", O_P_SSHIJI_K_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_P_SSHIJI_K]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_P_SSHIJI_K_Create()   '手配指図データ（子）作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "手配指図データ（子）マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "手配指図データ（子）マスタ")
                Exit Function
        End Select
    Loop
    
    O_P_SSHIJI_K_Open = False

End Function

