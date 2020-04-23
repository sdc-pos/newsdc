Attribute VB_Name = "SE_KOUTEI_TANKA_M"
Option Explicit
'********************************************************************
'*                                                                  *
'*              品目別作業工程単価設定マスタ  ファイル定義          *
'*                                                                  *
'*          CREATE 2008.02.05                                       *
'********************************************************************
'ファイルＩＤ
Public Const SE_KOUTEI_TANKA_M_ID$ = "SE_KOUTEI_TANKA_M"

'ページサイズ
Public Const SE_KOUTEI_TANKA_M_PG_SIZ% = 2048

'ポジション・ブロック
Public SE_KOUTEI_TANKA_M_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************

'前工程の構造体
Private Type MAE_KOUTEI_tag
    KOUSU(0 To 6)               As Byte     '工数 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '集計区分
    SEIKYU_SAKI(0 To 0)         As Byte     '請求先
End Type

'作業工程の構造体
Private Type SAGYO_KOUTEI_tag
    KOUTEI_NAME(0 To 39)        As Byte     '工程名
    KOUSU(0 To 6)               As Byte     '工数 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '集計区分
    SEIKYU_SAKI(0 To 0)         As Byte     '請求先
End Type

'後工程の構造体
Private Type ATO_KOUTEI_tag
    KOUSU(0 To 6)               As Byte     '工数 9(3)V999
    SYUKEI_KBN(0 To 0)          As Byte     '集計区分
    SEIKYU_SAKI(0 To 0)         As Byte     '請求先
End Type




'レコード定義
Type SE_KOUTEI_TANKA_M_REC_Tag
    
    SE_HIN_GAI(0 To 19)         As Byte     '品目コード
    
                                            '前工程
    SE_MAE_KOUTEI(0 To 9)       As MAE_KOUTEI_tag
                                            '作業工程
    SE_SAGYO_KOUTEI(0 To 19)    As SAGYO_KOUTEI_tag
                                            '後工程
    SE_ATO_KOUTEI(0 To 9)       As ATO_KOUTEI_tag
    
    FILLER(0 To 288)            As Byte
    
End Type
'データ・バッファ
Public SE_KOUTEI_TANKA_M_REC    As SE_KOUTEI_TANKA_M_REC_Tag

'キー定義

Type KEY0_SE_KOUTEI_TANKA_M                 'ＫＥＹ０
    SE_HIN_GAI(0 To 19)         As Byte     '品目コード
End Type
    
'キー・データ
Public K0_SE_KOUTEI_TANKA_M     As KEY0_SE_KOUTEI_TANKA_M

Type SE_KOUTEI_TANKA_M_FSpeck
    fs                  As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                 As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private SE_KOUTEI_TANKA_M_Speck As SE_KOUTEI_TANKA_M_FSpeck
Private Function SE_KOUTEI_TANKA_M_Create() As Integer
'********************************************************************
'*                                                                  *
'*              品目別作業工程単価設定マスタ  ＣＲＥＡＴＥ          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_KOUTEI_TANKA_M_Create = True
                                            '品目別作業工程単価設定マスタ   フルパス取込み
    sts = GetIni("FILE", SE_KOUTEI_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_KOUTEI_TANKA_M]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    SE_KOUTEI_TANKA_M_Speck.fs.recoleng = Len(SE_KOUTEI_TANKA_M_REC)    ' レコード長
    SE_KOUTEI_TANKA_M_Speck.fs.PageSize = SE_KOUTEI_TANKA_M_PG_SIZ      ' ページサイズ
    SE_KOUTEI_TANKA_M_Speck.fs.idexnumb = 1                             ' インデックス数
    SE_KOUTEI_TANKA_M_Speck.fs.fileflag = 0                             ' ファイルフラグ
    SE_KOUTEI_TANKA_M_Speck.fs.reserve = &H0                            ' 予約済み
    
    
    '-------------------------------------------'   ＫＥＹ０
    SE_KOUTEI_TANKA_M_Speck.ks0.keypos = 1                  ' キーポジション
    SE_KOUTEI_TANKA_M_Speck.ks0.keyleng = 20                ' キー長
    SE_KOUTEI_TANKA_M_Speck.ks0.keyflag = BtKfExt           ' キーフラグ
    SE_KOUTEI_TANKA_M_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SE_KOUTEI_TANKA_M_Speck.ks0.reserve = &H0               ' 予約済み
    '-------------------------------------------'   ＫＥＹ０

    sts = BTRV(BtOpCreate, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_Speck, Len(SE_KOUTEI_TANKA_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品目別作業工程単価設定マスタ")
        Exit Function
    End If
    
    SE_KOUTEI_TANKA_M_Create = False

End Function

Function SE_KOUTEI_TANKA_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目別作業工程単価設定マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    SE_KOUTEI_TANKA_M_Open = True
                                                '品目別作業工程単価設定マスタ   フルパス取込み
    sts = GetIni("FILE", SE_KOUTEI_TANKA_M_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_KOUTEI_TANKA_M]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_REC, Len(SE_KOUTEI_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_KOUTEI_TANKA_M_Create()    '品目別作業工程単価設定マスタ 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_KOUTEI_TANKA_M_POS, SE_KOUTEI_TANKA_M_REC, Len(SE_KOUTEI_TANKA_M_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品目別作業工程単価設定マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品目別作業工程単価設定マスタ")
                Exit Function
        End Select
    Loop
    SE_KOUTEI_TANKA_M_Open = False

End Function
