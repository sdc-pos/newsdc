Attribute VB_Name = "P_SSHIJI_K"
Option Explicit

'********************************************************************
'*
'*              商品化指図データ（子）  ファイル定義
'*
'*          CREATE 2005.11.11
'********************************************************************
'ファイルＩＤ


Public Const P_SSHIJI_K_ID$ = "P_SSHIJI_K"

'ページサイズ
Private Const P_SSHIJI_K_PG_SIZ% = 1024

'ポジション・ブロック
Public P_SSHIJI_K_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Public Type P_SSHIJI_K_REC_Tag
    
    xSHIJI_NO(0 To 4)        As Byte        '指図票�� 未使用とする 2007.11.28
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
    KO_SYUBETSU(0 To 1)     As Byte         '子　種別
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    KO_QTY(0 To 5)          As Byte         '子　員数(999V99)
    KO_SHIJI_QTY(0 To 10)   As Byte         '指示数(9(8)V99)
    KO_BIKOU(0 To 39)       As Byte         '子　備考
'    KO_ID_NO(0 To 7)        As Byte        '子 ＩＤ＿ＮＯ
    KO_ID_NO(0 To 11)       As Byte         '子 ＩＤ＿ＮＯ (8桁→12桁)  2006/05/24
    CALCEL_F(0 To 0)        As Byte         'ｷｬﾝｾﾙF
    CANCEL_DATETIME(0 To 13) As Byte        'ｷｬﾝｾﾙ日時
'    FILLER(0 To 64)         As Byte         'Filler
    
    SHIJI_No(0 To 7)        As Byte         '指図票��   2007.11.28
    
    
    HIKIATE_QTY(0 To 10)    As Byte         '在庫引当数 2012.03.09
    IDO_SUMI(0 To 0)        As Byte         '移動済み 空白:未　9:済み 2012.03.09
    
    ST_TANABAN(0 To 7)      As Byte         '標準棚番 2012.03.18
        
        
    IDO_SUMI_QTY(0 To 10)   As Byte         '移動済み数量 2012.04.13
        
        
        
    COMPO_TANTO(0 To 4)     As Byte         '構成ﾁｪｯｸ   担当者          2012.04.20
    COMPO_YMDHS(0 To 11)    As Byte         '           日時            2012.04.20
    COMPO_Sumi_Cnt(0 To 1)  As Byte         '           ﾁｪｯｸ済み数      2012.04.20
    COMPO_ALL_Cnt(0 To 1)   As Byte         '           構成数          2012.04.20
        
    FILLER(0 To 0)          As Byte         'Filler 2007.11.28  2012.04.20 桁数変更
    
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_SSHIJI_K_REC       As P_SSHIJI_K_REC_Tag

'キー定義

Type KEY0_P_SSHIJI_K                        'ＫＥＹ０
'    SHIJI_NO(0 To 4)        As Byte         '指図票��  2007.11.28
    SHIJI_No(0 To 7)        As Byte         '指図票��   '2007.11.28
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    SEQNO(0 To 2)           As Byte         '追番
End Type
    
    
Type KEY1_P_SSHIJI_K                        'ＫＥＹ１
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
'    KO_ID_NO(0 To 7)        As Byte         '子 ＩＤ＿ＮＯ
    KO_ID_NO(0 To 11)       As Byte         '子 ＩＤ＿ＮＯ (8桁→12桁)  2006/05/24
End Type
    
Type KEY2_P_SSHIJI_K                        'ＫＥＹ２   2012.03.09
    KO_JGYOBU(0 To 0)       As Byte         '子　事業部
    KO_NAIGAI(0 To 0)       As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)     As Byte         '子　品番
    IDO_SUMI(0 To 0)        As Byte         '移動済み 空白:未　9:済み
End Type
    
Type KEY3_P_SSHIJI_K                        'ＫＥＹ３   2012.03.18
    SHIJI_No(0 To 7)        As Byte         '指図票��
    DATA_KBN(0 To 0)        As Byte         'ﾃﾞｰﾀ区分
    ST_TANABAN(0 To 7)      As Byte         '標準棚番
End Type
    
    
    
    
'キー・データ
Public K0_P_SSHIJI_K        As KEY0_P_SSHIJI_K
Public K1_P_SSHIJI_K        As KEY1_P_SSHIJI_K
Public K2_P_SSHIJI_K        As KEY2_P_SSHIJI_K  '2012.03.09
Public K3_P_SSHIJI_K        As KEY3_P_SSHIJI_K  '2012.03.18

Type P_SSHIJI_K_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

    ks5                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.09
    ks6                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.09
    ks7                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.09
    ks8                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.09

    ks9                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.18
    ks10                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.18
    ks11                    As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体    2012.03.18


End Type

Private P_SSHIJI_K_Speck    As P_SSHIJI_K_FSpeck
Private Function P_SSHIJI_K_Create() As Integer
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

    P_SSHIJI_K_Create = True
                                            '手配指図データ（子）フルパス取込み
    sts = GetIni("FILE", P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_K]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SSHIJI_K_Speck.fs.recoleng = Len(P_SSHIJI_K_REC)  ' レコード長
    P_SSHIJI_K_Speck.fs.PageSize = P_SSHIJI_K_PG_SIZ    ' ページサイズ
    P_SSHIJI_K_Speck.fs.idexnumb = 4                    ' インデックス数
    P_SSHIJI_K_Speck.fs.fileflag = 0                    ' ファイルフラグ
    P_SSHIJI_K_Speck.fs.reserve = &H0                   ' 予約済み
    '--------------------------------------------------- キー０ ▽
'2008.11.28    P_SSHIJI_K_Speck.ks0.keypos = 1                     ' キーポジション
'2008.11.28    P_SSHIJI_K_Speck.ks0.keyleng = 5                    ' キー長
    
    P_SSHIJI_K_Speck.ks0.keypos = 118                   ' キーポジション
    P_SSHIJI_K_Speck.ks0.keyleng = 8                    ' キー長
    
    
    P_SSHIJI_K_Speck.ks0.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    P_SSHIJI_K_Speck.ks0.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks0.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_K_Speck.ks1.keypos = 6                     ' キーポジション
    P_SSHIJI_K_Speck.ks1.keyleng = 1                    ' キー長
    P_SSHIJI_K_Speck.ks1.keyflag = BtKfExt + BtKfSeg    ' キーフラグ
    P_SSHIJI_K_Speck.ks1.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks1.reserve = &H0                  ' 予約済み
    
    
    P_SSHIJI_K_Speck.ks2.keypos = 7                     ' キーポジション
    P_SSHIJI_K_Speck.ks2.keyleng = 3                    ' キー長
    P_SSHIJI_K_Speck.ks2.keyflag = BtKfExt              ' キーフラグ
    P_SSHIJI_K_Speck.ks2.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks2.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー０ △
    
    '--------------------------------------------------- キー１ ▽
    P_SSHIJI_K_Speck.ks3.keypos = 12                    ' キーポジション
    P_SSHIJI_K_Speck.ks3.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfSeg
    P_SSHIJI_K_Speck.ks3.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks3.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_K_Speck.ks4.keypos = 91                    ' キーポジション
    P_SSHIJI_K_Speck.ks4.keyleng = 12                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks4.keyflag = BtKfExt + BtKfDup
    P_SSHIJI_K_Speck.ks4.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks4.reserve = &H0                  ' 予約済み
    '--------------------------------------------------- キー１ △
    
    
    '--------------------------------------------------- キー２ ▽  2012.03.09
    P_SSHIJI_K_Speck.ks5.keypos = 12                    ' キーポジション
    P_SSHIJI_K_Speck.ks5.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks5.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks5.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks5.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_K_Speck.ks6.keypos = 13                    ' キーポジション
    P_SSHIJI_K_Speck.ks6.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks6.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks6.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks6.reserve = &H0                  ' 予約済み
    
    
    P_SSHIJI_K_Speck.ks7.keypos = 14                    ' キーポジション
    P_SSHIJI_K_Speck.ks7.keyleng = 20                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks7.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks7.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks7.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_K_Speck.ks8.keypos = 137                   ' キーポジション
    P_SSHIJI_K_Speck.ks8.keyleng = 1                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks8.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SSHIJI_K_Speck.ks8.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks8.reserve = &H0                  ' 予約済み
    
    '--------------------------------------------------- キー２ △
    
    '--------------------------------------------------- キー３  ▽  2012.03.09
    P_SSHIJI_K_Speck.ks9.keypos = 118                   ' キーポジション
    P_SSHIJI_K_Speck.ks9.keyleng = 8                    ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks9.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks9.keytype = Chr(BtKtString)      ' キータイプ
    P_SSHIJI_K_Speck.ks9.reserve = &H0                  ' 予約済み
    
    P_SSHIJI_K_Speck.ks10.keypos = 6                    ' キーポジション
    P_SSHIJI_K_Speck.ks10.keyleng = 1                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks10.keyflag = BtKfExt + BtKfDup + BtKfChg + BtKfSeg
    P_SSHIJI_K_Speck.ks10.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_K_Speck.ks10.reserve = &H0                 ' 予約済み
    
    
    P_SSHIJI_K_Speck.ks11.keypos = 138                  ' キーポジション
    P_SSHIJI_K_Speck.ks11.keyleng = 8                   ' キー長
                                                        ' キーフラグ
    P_SSHIJI_K_Speck.ks11.keyflag = BtKfExt + BtKfDup + BtKfChg
    P_SSHIJI_K_Speck.ks11.keytype = Chr(BtKtString)     ' キータイプ
    P_SSHIJI_K_Speck.ks11.reserve = &H0                 ' 予約済み
    '--------------------------------------------------- キー３ △
    
    
    sts = BTRV(BtOpCreate, P_SSHIJI_K_POS, P_SSHIJI_K_Speck, Len(P_SSHIJI_K_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "手配指図データ（子）")
        Exit Function
    End If
    
    P_SSHIJI_K_Create = False

End Function

Public Function P_SSHIJI_K_Open(Mode As Integer) As Integer
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

    P_SSHIJI_K_Open = True
                                            '手配指図データ（子）フルパス取込み
    sts = GetIni("FILE", P_SSHIJI_K_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_K]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SSHIJI_K_Create()   '手配指図データ（子）作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), ByVal FullPath, Len(FullPath), Mode)
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
    
    P_SSHIJI_K_Open = False

End Function
