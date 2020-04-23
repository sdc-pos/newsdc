Attribute VB_Name = "OYA_ITEM"
Option Explicit
'********************************************************************
'*                                                                  *
'*              親部品展開データ　ファイル定義                      *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
'ファイルＩＤ
Public Const OYA_ITEM_ID$ = "OYA_ITEM"

'ページサイズ
Public Const OYA_ITEM_PG_SIZ% = 2048

'ポジション・ブロック
Public OYA_ITEM_POS         As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type OYA_ITEMREC_Tag
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）

    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数

    ST_SOKO(0 To 1)         As Byte     '標準入庫倉庫 倉庫
    ST_RETU(0 To 1)         As Byte     '             列
    ST_REN(0 To 1)          As Byte     '             連
    ST_DAN(0 To 1)          As Byte     '             段




End Type

'データ・バッファ
Public OYA_ITEMREC          As OYA_ITEMREC_Tag

'キー定義
Private Type KEY0_OYA_ITEM          'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

Private Type KEY1_OYA_ITEM          'ＫＥＹ１
    AVE_SYUKA(0 To 7)       As Byte     '平均出荷数
    JGYOBU(0 To 0)          As Byte     '事業部区分
    NAIGAI(0 To 0)          As Byte     '国内外
    HIN_GAI(0 To 19)        As Byte     '品番（外部）
End Type

'キー・データ
Public K0_OYA_ITEM          As KEY0_OYA_ITEM
Public K1_OYA_ITEM          As KEY1_OYA_ITEM

Private Type OYA_ITEM_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks3 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks4 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks5 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks6 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private OYA_ITEM_Speck      As OYA_ITEM_FSpeck

Private Function OYA_ITEM_Create() As Integer
'********************************************************************
'*                                                                  *
'*              親部品展開データ　ＣＲＥＡＴＥ                      *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128


Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer




    OYA_ITEM_Create = True
                                            '在庫集計データフルパス取込み
    sts = GetIni("FILE", OYA_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[OYA_ITEM] 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If


    Ret = InStr(1, FullPath, ".") - 1
    FullPath = Left(FullPath, Ret) & com & Right(FullPath, Len(FullPath) - Ret)
    
    
    
    
    
    OYA_ITEM_Speck.fs.recoleng = Len(OYA_ITEMREC)       ' レコード長
    OYA_ITEM_Speck.fs.PageSize = OYA_ITEM_PG_SIZ        ' ページサイズ
    OYA_ITEM_Speck.fs.idexnumb = 2                      ' インデックス数
    OYA_ITEM_Speck.fs.fileflag = 0                      ' ファイルフラグ
    OYA_ITEM_Speck.fs.reserve = &H0                     ' 予約済み
'-----------------------------------------------' キー０
    OYA_ITEM_Speck.ks0.keypos = 1                       ' キーポジション
    OYA_ITEM_Speck.ks0.keyleng = 1                      ' キー長
    OYA_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    OYA_ITEM_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks0.reserve = &H0                    ' 予約済み

    OYA_ITEM_Speck.ks1.keypos = 2                       ' キーポジション
    OYA_ITEM_Speck.ks1.keyleng = 1                      ' キー長
    OYA_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    OYA_ITEM_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks1.reserve = &H0                    ' 予約済み

    OYA_ITEM_Speck.ks2.keypos = 3                       ' キーポジション
    OYA_ITEM_Speck.ks2.keyleng = 20                     ' キー長
    OYA_ITEM_Speck.ks2.keyflag = BtKfExt                ' キーフラグ
    OYA_ITEM_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks2.reserve = &H0                    ' 予約済み
'-----------------------------------------------' キー１

    OYA_ITEM_Speck.ks3.keypos = 23                      ' キーポジション
    OYA_ITEM_Speck.ks3.keyleng = 8                      ' キー長
    OYA_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    OYA_ITEM_Speck.ks3.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks3.reserve = &H0                    ' 予約済み
    
    OYA_ITEM_Speck.ks4.keypos = 1                       ' キーポジション
    OYA_ITEM_Speck.ks4.keyleng = 1                      ' キー長
    OYA_ITEM_Speck.ks4.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    OYA_ITEM_Speck.ks4.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks4.reserve = &H0                    ' 予約済み

    OYA_ITEM_Speck.ks5.keypos = 2                       ' キーポジション
    OYA_ITEM_Speck.ks5.keyleng = 1                      ' キー長
    OYA_ITEM_Speck.ks5.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    OYA_ITEM_Speck.ks5.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks5.reserve = &H0                    ' 予約済み

    OYA_ITEM_Speck.ks6.keypos = 3                       ' キーポジション
    OYA_ITEM_Speck.ks6.keyleng = 20                     ' キー長
    OYA_ITEM_Speck.ks6.keyflag = BtKfExt                ' キーフラグ
    OYA_ITEM_Speck.ks6.keytype = Chr(BtKtString)        ' キータイプ
    OYA_ITEM_Speck.ks6.reserve = &H0                    ' 予約済み

    sts = BTRV(BtOpCreate, OYA_ITEM_POS, OYA_ITEM_Speck, Len(OYA_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "親部品展開データ")
        Exit Function
    End If
    
    OYA_ITEM_Create = False

End Function

Function OYA_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              親部品展開データ　ＯＰＥＮ                          *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 2008.11.05                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    
Dim sBuffer     As String * 255
Dim com         As String


Dim Ret         As Integer
    
    
    OYA_ITEM_Open = True
                                            '在庫集計データフルパス取込み
    sts = GetIni("FILE", OYA_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[OYA_ITEM] 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    
    On Error Resume Next
    
    
    Kill (FullPath)
    
    On Error GoTo 0
    
    
    
    
    
    Do
        sts = BTRV(BtOpOpen, OYA_ITEM_POS, OYA_ITEMREC, Len(SUMZREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = OYA_ITEM_Create()        '在庫集計データ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, OYA_ITEM_POS, OYA_ITEMREC, Len(OYA_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "親部品展開データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "親部品展開データ")
                Exit Function
        End Select
    Loop

    OYA_ITEM_Open = False
End Function


