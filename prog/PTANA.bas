Attribute VB_Name = "PTANA"
Option Explicit
'********************************************************************
'*                                                                  *
'*              個装箱用棚リスト印刷ファイル（一時ファイル）        *
'*                                                                  *
'*          CREATE 2004.04.23                                       *
'********************************************************************
'ファイルＩＤ
Public Const PTANA_ID = "PTANA"

'ページサイズ
Public Const PTANA_PG_SIZ% = 512

'ポジション・ブロック
Public PTANA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type PTANAREC_Tag
    Packing_No(0 To 3)  As Byte     '個装箱№
    Rank(0 To 2)        As Byte     'ランク
    Page_cnt(0 To 0)    As Byte     'ページ数(倉庫毎)
    SEQ(0 To 4)         As Byte     'SEQ番号
    SOKO_NO01(0 To 1)   As Byte     '倉庫１
    RETUREN01(0 To 4)   As Byte     '列・連１
    SOKO_NO02(0 To 1)   As Byte     '倉庫２
    RETUREN02(0 To 4)   As Byte     '列・連２
    SOKO_NO03(0 To 1)   As Byte     '倉庫３
    RETUREN03(0 To 4)   As Byte     '列・連３
    SOKO_NO04(0 To 1)   As Byte     '倉庫４
    RETUREN04(0 To 4)   As Byte     '列・連４
    SOKO_NO05(0 To 1)   As Byte     '倉庫５
    RETUREN05(0 To 4)   As Byte     '列・連５
    SOKO_NO06(0 To 1)   As Byte     '倉庫６
    RETUREN06(0 To 4)   As Byte     '列・連６
    SOKO_NO07(0 To 1)   As Byte     '倉庫７
    RETUREN07(0 To 4)   As Byte     '列・連７
    SOKO_NO08(0 To 1)   As Byte     '倉庫８
    RETUREN08(0 To 4)   As Byte     '列・連８
    SOKO_NO09(0 To 1)   As Byte     '倉庫９
    RETUREN09(0 To 4)   As Byte     '列・連９
    SOKO_NO10(0 To 1)   As Byte     '倉庫１０
    RETUREN10(0 To 4)   As Byte     '列・連１０

End Type
'データ・バッファ
Public PTANAREC         As PTANAREC_Tag


'キー定義
Type KEY0_PTANA                     'ＫＥＹ０
    Packing_No(0 To 3)  As Byte     '個装箱№
    Rank(0 To 2)        As Byte     'ランク
    Page_cnt(0 To 0)    As Byte     'ページ数(倉庫毎)
    SEQ(0 To 4)         As Byte     'SEQ番号
End Type
    
'キー・データ
Public K0_PTANA         As KEY0_PTANA

Private Type PTANA_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private PTANA_Speck    As PTANA_FSpeck
Private Function PTANA_Create() As Integer
'********************************************************************
'*                                                                  *
'*              個装箱別棚リスト印刷ファイル  ＣＲＥＡＴＥ          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.24                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PTANA_Create = True
                                            '個装箱別棚リスト印刷ファイルフルパス取込み
    sts = GetIni("FILE", PTANA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PTANA] 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PTANA_Speck.fs.recoleng = Len(PTANAREC)         ' レコード長
    PTANA_Speck.fs.PageSize = PTANA_PG_SIZ          ' ページサイズ
    PTANA_Speck.fs.idexnumb = 1                     ' インデックス数
    PTANA_Speck.fs.fileflag = 0                     ' ファイルフラグ
    PTANA_Speck.fs.reserve = &H0                    ' 予約済み
                                                    
'---------------------------------------------------' キー０
    PTANA_Speck.ks0.keypos = 1                      ' キーポジション
    PTANA_Speck.ks0.keyleng = 4                     ' キー長
    PTANA_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    PTANA_Speck.ks0.keytype = Chr(BtKtString)       ' キータイプ
    PTANA_Speck.ks0.reserve = &H0                   ' 予約済み

    PTANA_Speck.ks1.keypos = 5                      ' キーポジション
    PTANA_Speck.ks1.keyleng = 3                     ' キー長
    PTANA_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    PTANA_Speck.ks1.keytype = Chr(BtKtString)       ' キータイプ
    PTANA_Speck.ks1.reserve = &H0                   ' 予約済み

    PTANA_Speck.ks2.keypos = 8                      ' キーポジション
    PTANA_Speck.ks2.keyleng = 1                     ' キー長
    PTANA_Speck.ks2.keyflag = BtKfExt + BtKfSeg     ' キーフラグ
    PTANA_Speck.ks2.keytype = Chr(BtKtString)       ' キータイプ
    PTANA_Speck.ks2.reserve = &H0                   ' 予約済み

    PTANA_Speck.ks3.keypos = 9                      ' キーポジション
    PTANA_Speck.ks3.keyleng = 5                     ' キー長
    PTANA_Speck.ks3.keyflag = BtKfExt               ' キーフラグ
    PTANA_Speck.ks3.keytype = Chr(BtKtString)       ' キータイプ
    PTANA_Speck.ks3.reserve = &H0                   ' 予約済み

    sts = BTRV(BtOpCreate, PTANA_POS, PTANA_Speck, Len(PTANA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "個装箱別棚リスト印刷ファイル")
        Exit Function
    End If
    
    PTANA_Create = False

End Function

Function PTANA_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              個装箱別棚リスト印刷ファイル  ＯＰＥＮ              *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.04.24                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PTANA_Open = True
                                            '個装箱別棚リスト印刷ファイルフルパス取込み
    sts = GetIni("FILE", PTANA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PTANA_POS, PTANAREC, Len(PTANAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PTANA_Create()        '個装箱別棚リスト印刷ファイル作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PTANA_POS, PTANAREC, Len(PTANAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "個装箱別棚リスト印刷ファイル")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "個装箱別棚リスト印刷ファイル")
                Exit Function
        End Select
    Loop
    PTANA_Open = False
End Function
