Attribute VB_Name = "SE_SHOHIN_TANKA_M"
Option Explicit
'********************************************************************
'*                                                                  *
'*              商品化単価マスタ  ファイル定義                      *
'*                                                                  *
'*          CREATE 2008.02.05                                       *
'********************************************************************
'ファイルＩＤ
Public Const SE_SHOHIN_TANKA_M_ID$ = "SE_SHOHIN_TANKA_M"

'ページサイズ
Public Const SE_SHOHIN_TANKA_M_PG_SIZ% = 512

'ポジション・ブロック
Public SE_SHOHIN_TANKA_M_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SE_SHOHIN_TANKA_M_REC_Tag
    
    SE_HIN_GAI(0 To 19)         As Byte     'ＫＥＹ
    
    SE_KOU_TANKA(0 To 10)       As Byte     '工数　単価 9(8)V99
    SE_KOU_SET_DATE(0 To 7)     As Byte     '工数　単価設定日

    SE_SIZ_TANKA(0 To 10)       As Byte     '資材　単価 9(8)V99
    SE_SIZ_SET_DATE(0 To 7)     As Byte     '資材　単価設定日


    FILLER(0 To 198)            As Byte
    
End Type
'データ・バッファ
Public SE_SHOHIN_TANKA_M_REC    As SE_SHOHIN_TANKA_M_REC_Tag

'キー定義

Type KEY0_SE_SHOHIN_TANKA_M                 'ＫＥＹ０
    SE_HIN_GAI(0 To 19)         As Byte     'ＫＥＹ
End Type
    
'キー・データ
Public K0_SE_SHOHIN_TANKA_M     As KEY0_SE_SHOHIN_TANKA_M

Type SE_SHOHIN_TANKA_M_FSpeck
    fs                  As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SE_SHOHIN_TANKA_M_Speck As SE_SHOHIN_TANKA_M_FSpeck
Private Function SE_SHOHIN_TANKA_M_Create() As Integer
'********************************************************************
'*                                                                  *
'*              商品化単価マスタ  ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_SHOHIN_TANKA_M_Create = True
                                            '商品化単価マスタ   フルパス取込み
    sts = GetIni("FILE", SE_SHOHIN_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_SHOHIN_TANKA_M]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    SE_SHOHIN_TANKA_M_Speck.fs.recoleng = Len(SE_SHOHIN_TANKA_M_REC)    ' レコード長
    SE_SHOHIN_TANKA_M_Speck.fs.PageSize = SE_SHOHIN_TANKA_M_PG_SIZ      ' ページサイズ
    SE_SHOHIN_TANKA_M_Speck.fs.idexnumb = 1                             ' インデックス数
    SE_SHOHIN_TANKA_M_Speck.fs.fileflag = 0                             ' ファイルフラグ
    SE_SHOHIN_TANKA_M_Speck.fs.reserve = &H0                            ' 予約済み
    
    
    '-------------------------------------------'   ＫＥＹ０
    SE_SHOHIN_TANKA_M_Speck.ks0.keypos = 1                 ' キーポジション
    SE_SHOHIN_TANKA_M_Speck.ks0.keyleng = 2                ' キー長
    SE_SHOHIN_TANKA_M_Speck.ks0.keyflag = BtKfExt          ' キーフラグ
    SE_SHOHIN_TANKA_M_Speck.ks0.keytype = Chr(BtKtString)  ' キータイプ
    SE_SHOHIN_TANKA_M_Speck.ks0.reserve = &H0              ' 予約済み
    '-------------------------------------------'   ＫＥＹ０

    sts = BTRV(BtOpCreate, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_Speck, Len(SE_SHOHIN_TANKA_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ロケーション別単価設定マスタ")
        Exit Function
    End If
    
    SE_SHOHIN_TANKA_M_Create = False

End Function

Function SE_SHOHIN_TANKA_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ロケーション別単価設定マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SE_SHOHIN_TANKA_M_Open = True
                                                'ロケーション別単価設定マスタ   フルパス取込み
    sts = GetIni("FILE", SE_SHOHIN_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_SHOHIN_TANKA_M]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_REC, Len(SE_SHOHIN_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_SHOHIN_TANKA_M_Create()   'ロケーション別単価設定マスタ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_SHOHIN_TANKA_M_POS, SE_SHOHIN_TANKA_M_REC, Len(SE_SHOHIN_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ロケーション別単価設定マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ロケーション別単価設定マスタ")
                Exit Function
        End Select
    Loop
    SE_SHOHIN_TANKA_M_Open = False

End Function
