Attribute VB_Name = "B_ITEM"
Option Explicit
'********************************************************************
'*
'*              美的品番管理データ  ファイル定義
'*
'*          CREATE 2013.10.17
'********************************************************************
'ファイルＩＤ
Public Const B_ITEM_ID$ = "B_ITEM"

'ページサイズ
Public Const B_ITEM_PG_SIZ% = 1024

'ポジション・ブロック
Public B_ITEM_POS         As POSBLK
'=
'====================================================================
'=          レコード初期化プロシージャ     ( Rclr_ITEMREC )
'====================================================================
'=
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************

'レコード定義
Type B_ITEMREC_Tag
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    
    B_HIN_CODE(0 To 69)         As Byte     '美的品番ｺｰﾄﾞ
    
    FILLER(0 To 371)            As Byte     'FILLER
    

    INS_TANTO(0 To 9)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時
    UPD_TANTO(0 To 9)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type
'データ・バッファ
Public B_ITEMREC                As B_ITEMREC_Tag

'キー定義

Type KEY0_B_ITEM                    'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type


Type KEY1_B_ITEM                    'ＫＥＹ１
    B_HIN_CODE(0 To 69)         As Byte     '美的品番ｺｰﾄﾞ
End Type


'キー・データ
Public K0_B_ITEM                As KEY0_B_ITEM
Public K1_B_ITEM                As KEY1_B_ITEM

Type B_ITEM_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
End Type

Private B_ITEM_Speck    As B_ITEM_FSpeck

Private Function B_ITEMreate() As Integer
'********************************************************************
'*
'*              美的品番管理データ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    B_ITEMreate = True
                                            '美的品番管理データ フルパス取込み
    sts = GetIni("FILE", B_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [B_ITEM]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    B_ITEM_Speck.fs.recoleng = Len(B_ITEMREC)   ' レコード長
    B_ITEM_Speck.fs.PageSize = B_ITEM_PG_SIZ    ' ページサイズ
    B_ITEM_Speck.fs.idexnumb = 2                ' インデックス数
    B_ITEM_Speck.fs.fileflag = 0                ' ファイルフラグ
    B_ITEM_Speck.fs.reserve = &H0               ' 予約済み
'-----------------------------------------------
                                                ' キー０
    B_ITEM_Speck.ks0.keypos = 1                             ' キーポジション
    B_ITEM_Speck.ks0.keyleng = 1                            ' キー長
    B_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' キーフラグ
    B_ITEM_Speck.ks0.keytype = Chr(BtKtString)              ' キータイプ
    B_ITEM_Speck.ks0.reserve = &H0                          ' 予約済み

    B_ITEM_Speck.ks1.keypos = 2                             ' キーポジション
    B_ITEM_Speck.ks1.keyleng = 1                            ' キー長
    B_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfSeg  ' キーフラグ
    B_ITEM_Speck.ks1.keytype = Chr(BtKtString)              ' キータイプ
    B_ITEM_Speck.ks1.reserve = &H0                          ' 予約済み

    B_ITEM_Speck.ks2.keypos = 3                             ' キーポジション
    B_ITEM_Speck.ks2.keyleng = 20                           ' キー長
    B_ITEM_Speck.ks2.keyflag = BtKfExt + BtKfChg            ' キーフラグ
    B_ITEM_Speck.ks2.keytype = Chr(BtKtString)              ' キータイプ
    B_ITEM_Speck.ks2.reserve = &H0                          ' 予約済み
'-----------------------------------------------

'-----------------------------------------------
                                                ' キー１
    B_ITEM_Speck.ks3.keypos = 23                            ' キーポジション
    B_ITEM_Speck.ks3.keyleng = 70                           ' キー長
    B_ITEM_Speck.ks3.keyflag = BtKfExt + BtKfDup + BtKfChg  ' キーフラグ
    B_ITEM_Speck.ks3.keytype = Chr(BtKtString)              ' キータイプ
    B_ITEM_Speck.ks3.reserve = &H0                          ' 予約済み
'-----------------------------------------------



    sts = BTRV(BtOpCreate, B_ITEM_POS, B_ITEM_Speck, Len(B_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "美的品番管理ﾃﾞｰﾀ")
        Exit Function
    End If

    B_ITEMreate = False

End Function

Public Function B_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              美的品番管理データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    B_ITEM_Open = True
                                                '美的品番管理データ フルパス取込み
    sts = GetIni("FILE", B_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [B_ITEM]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = B_ITEMreate()             '美的品番管理データ    作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "美的品番管理データ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "美的品番管理データ")
                Exit Function
        End Select
    Loop

    B_ITEM_Open = False

End Function

Public Sub Rclr_B_ITEMREC()
'********************************************************************
'*
'*              美的品番管理データ  レコード初期化
'*
'********************************************************************


    Call UniCode_Conv(B_ITEMREC.JGYOBU, "")             '事業部区分
    Call UniCode_Conv(B_ITEMREC.NAIGAI, "")             '国内外
    Call UniCode_Conv(B_ITEMREC.HIN_GAI, "")            '品番（外部）


    Call UniCode_Conv(B_ITEMREC.B_HIN_CODE, "")         '美的品番
    
    Call UniCode_Conv(B_ITEMREC.FILLER, "")

    Call UniCode_Conv(B_ITEMREC.INS_TANTO, "")          '追加担当
    Call UniCode_Conv(B_ITEMREC.Ins_DateTime, "")       '追加日時

    Call UniCode_Conv(B_ITEMREC.UPD_TANTO, "")          '更新担当
    Call UniCode_Conv(B_ITEMREC.UPD_DATETIME, "")       '更新日時

End Sub
