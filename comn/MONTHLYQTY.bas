Attribute VB_Name = "MONTHLYQTY"
Option Explicit
'********************************************************************
'*                                                                  *
'*              月平均出荷数(月別集計)  ファイル定義                *
'*                                                                  *
'*          CREATE 2008.07.08                                       *
'********************************************************************
'ファイルＩＤ
Public Const MONTHLYQTY_ID$ = "MONTHLYQTY"

'ページサイズ
Public Const MONTHLYQTY_PG_SIZ% = 512

'ポジション・ブロック
Public MONTHLYQTY_POS       As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type MONTHLYQTYREC_Tag
    DT(0 To 7)              As Byte         '日付
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番(外部)
    SyukaCnt(0 To 4)        As Byte         '出荷回数
    SyukaQty(0 To 4)        As Byte         '出荷数量



End Type

'データ・バッファ
Public MONTHLYQTYREC        As MONTHLYQTYREC_Tag

'キー定義

Type KEY0_MONTHLYQTY                    'ＫＥＹ０
    DT(0 To 7)              As Byte         '日付
    JGYOBU(0 To 0)          As Byte         '事業部
    NAIGAI(0 To 0)          As Byte         '国内外
    HIN_GAI(0 To 19)        As Byte         '品番(外部)
End Type

'Type KEY1_MONTHLYQTY                    'ＫＥＹ１
'    JGYOBU(0 To 0)          As Byte         '事業部
'    NAIGAI(0 To 0)          As Byte         '国内外
'    HIN_GAI(0 To 19)        As Byte         '品番(外部)
'    DT(0 To 7)              As Byte         '日付
'End Type

'キー・データ
Public K0_MONTHLYQTY        As KEY0_MONTHLYQTY
'Public K1_MONTHLYQTY        As KEY1_MONTHLYQTY

Type MONTHLYQTY_FSpeck
    fs  As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
'    ks4 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
'    ks5 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
'    ks6 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
'    ks7 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Public MONTHLYQTY_Speck     As MONTHLYQTY_FSpeck
 
Private Function MONTHLYQTY_Create() As Integer
'********************************************************************
'*                                                                  *
'*              月平均出荷数(月別集計)  ＣＲＥＡＴＥ                      　  *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.14                                       *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    MONTHLYQTY_Create = True
                                            '月平均出荷数(月別集計)フルパス取込み
    sts = GetIni("FILE", MONTHLYQTY_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MONTHLYQTY_Speck.fs.recoleng = Len(MONTHLYQTYREC)   ' レコード長
    MONTHLYQTY_Speck.fs.PageSize = MONTHLYQTY_PG_SIZ%   ' ページサイズ
    MONTHLYQTY_Speck.fs.idexnumb = 1                    ' インデックス数
    MONTHLYQTY_Speck.fs.fileflag = 0                    ' ファイルフラグ
    MONTHLYQTY_Speck.fs.reserve = &H0                   ' 予約済み
                                                        ' キー０
    MONTHLYQTY_Speck.ks0.keypos = 1                         ' キーポジション
    MONTHLYQTY_Speck.ks0.keyleng = 8                        ' キー長
    MONTHLYQTY_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    MONTHLYQTY_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    MONTHLYQTY_Speck.ks0.reserve = &H0                      ' 予約済み
                                                        ' キー０
    MONTHLYQTY_Speck.ks1.keypos = 9                         ' キーポジション
    MONTHLYQTY_Speck.ks1.keyleng = 1                        ' キー長
    MONTHLYQTY_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    MONTHLYQTY_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    MONTHLYQTY_Speck.ks1.reserve = &H0                      ' 予約済み
                                                        ' キー０
    MONTHLYQTY_Speck.ks2.keypos = 10                        ' キーポジション
    MONTHLYQTY_Speck.ks2.keyleng = 1                        ' キー長
    MONTHLYQTY_Speck.ks2.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    MONTHLYQTY_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    MONTHLYQTY_Speck.ks2.reserve = &H0                      ' 予約済み
                                                        ' キー０
    MONTHLYQTY_Speck.ks3.keypos = 11                        ' キーポジション
    MONTHLYQTY_Speck.ks3.keyleng = 20                       ' キー長
    MONTHLYQTY_Speck.ks3.keyflag = BtKfExt                  ' キーフラグ
    MONTHLYQTY_Speck.ks3.keytype = Chr(BtKtString)          ' キータイプ
    MONTHLYQTY_Speck.ks3.reserve = &H0                      ' 予約済み




'                                                        ' キー１
'    MONTHLYQTY_Speck.ks4.keypos = 9                         ' キーポジション
'    MONTHLYQTY_Speck.ks4.keyleng = 1                        ' キー長
'    MONTHLYQTY_Speck.ks4.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
'    MONTHLYQTY_Speck.ks4.keytype = Chr(BtKtString)          ' キータイプ
'    MONTHLYQTY_Speck.ks4.reserve = &H0                      ' 予約済み
'                                                        ' キー１
'    MONTHLYQTY_Speck.ks5.keypos = 10                        ' キーポジション
'    MONTHLYQTY_Speck.ks5.keyleng = 1                        ' キー長
'    MONTHLYQTY_Speck.ks5.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
'    MONTHLYQTY_Speck.ks5.keytype = Chr(BtKtString)          ' キータイプ
'    MONTHLYQTY_Speck.ks5.reserve = &H0                      ' 予約済み
'                                                        ' キー１
'    MONTHLYQTY_Speck.ks6.keypos = 11                        ' キーポジション
'    MONTHLYQTY_Speck.ks6.keyleng = 20                       ' キー長
'    MONTHLYQTY_Speck.ks6.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
'    MONTHLYQTY_Speck.ks6.keytype = Chr(BtKtString)          ' キータイプ
'    MONTHLYQTY_Speck.ks6.reserve = &H0                      ' 予約済み
'                                                        ' キー１
'    MONTHLYQTY_Speck.ks7.keypos = 1                         ' キーポジション
'    MONTHLYQTY_Speck.ks7.keyleng = 8                        ' キー長
'    MONTHLYQTY_Speck.ks7.keyflag = BtKfExt                  ' キーフラグ
'    MONTHLYQTY_Speck.ks7.keytype = Chr(BtKtString)          ' キータイプ
'    MONTHLYQTY_Speck.ks7.reserve = &H0                      ' 予約済み



    sts = BTRV(BtOpCreate, MONTHLYQTY_POS, MONTHLYQTY_Speck, Len(MONTHLYQTY_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "月平均出荷数(月別集計)")
    End If

    MONTHLYQTY_Create = False

End Function

Function MONTHLYQTY_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              月平均出荷数(月別集計)  ＯＰＥＮ                    *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2008.07.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    MONTHLYQTY_Open = True
                                            '月平均出荷数(月別集計) フルパス取込み
    sts = GetIni("FILE", MONTHLYQTY_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MONTHLYQTY_Create()   '月平均出荷数(月別集計) 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "月平均出荷数(月別集計)")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "月平均出荷数(月別集計)")
                Exit Function
        End Select
    Loop

    MONTHLYQTY_Open = False
    
End Function
