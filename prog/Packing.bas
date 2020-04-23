Attribute VB_Name = "PACKING"
Option Explicit
'********************************************************************
'*                                                                  *
'*              個装箱マスタ  ファイル定義                          *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
'ファイルＩＤ
Public Const PACKING_ID = "PACKING"

'ページサイズ
Public Const PACKING_PG_SIZ% = 512

'ポジション・ブロック
Public PACKING_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type PACKINGREC_Tag
    PACKING_NO(0 To 3)  As Byte         '個装箱№
    RANK_A1(0 To 7)     As Byte         'ランク　Ａ－１
    RANK_A2(0 To 7)     As Byte         'ランク　Ａ－２
    RANK_B1(0 To 7)     As Byte         'ランク　Ｂ－１
    RANK_B2(0 To 7)     As Byte         'ランク　Ｂ－２
    RANK_C1(0 To 7)     As Byte         'ランク　Ｃ－１
    RANK_C2(0 To 7)     As Byte         'ランク　Ｃ－２
    FILLER(0 To 43)     As Byte         'FILLER
End Type
'データ・バッファ
Public PACKINGREC       As PACKINGREC_Tag


'キー定義
Type KEY0_PACKING                       'ＫＥＹ０
    PACKING_NO(0 To 3)  As Byte         '個装箱№
End Type
    
'キー・データ
Public K0_PACKING       As KEY0_PACKING

Type PACKING_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private PACKING_Speck    As PACKING_FSpeck
Private Function PACKING_Create() As Integer
'********************************************************************
'*                                                                  *
'*              個装箱マスタ  ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PACKING_Create = True
                                            '個装箱マスタフルパス取込み
    sts = GetIni("FILE", PACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PACKING_Speck.fs.recoleng = Len(PACKINGREC)     ' レコード長
    PACKING_Speck.fs.PageSize = PACKING_PG_SIZ      ' ページサイズ
    PACKING_Speck.fs.idexnumb = 1                   ' インデックス数
    PACKING_Speck.fs.fileflag = 0                   ' ファイルフラグ
    PACKING_Speck.fs.reserve = &H0                  ' 予約済み
                                                    ' キー０
    PACKING_Speck.ks0.keypos = 1                    ' キーポジション
    PACKING_Speck.ks0.keyleng = 4                   ' キー長
    PACKING_Speck.ks0.keyflag = BtKfExt             ' キーフラグ
    PACKING_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    PACKING_Speck.ks0.reserve = &H0                 ' 予約済み

    sts = BTRV(BtOpCreate, PACKING_POS, PACKING_Speck, Len(PACKING_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "個装箱マスタ")
        Exit Function
    End If

    PACKING_Create = False

End Function

Function PACKING_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              個装箱マスタ  ＯＰＥＮ                              *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PACKING_Open = True
                                            '個装箱マスタフルパス取込み
    sts = GetIni("FILE", PACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PACKING_POS, PACKINGREC, Len(PACKINGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PACKING_Create()        '個装箱マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PACKING_POS, PACKINGREC, Len(PACKINGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "個装箱マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "個装箱マスタ")
                Exit Function
        End Select
    Loop
    PACKING_Open = False
End Function
