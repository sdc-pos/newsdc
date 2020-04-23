Attribute VB_Name = "ITEM_LOC"
Option Explicit
'********************************************************************
'*
'*              品目－棚マスタ  ファイル定義
'*
'*          CREATE 2012.06.01
'********************************************************************
'ファイルＩＤ
Public Const ITEM_LOC_ID$ = "ITEM_LOC"

'ページサイズ
Public Const ITEM_LOC_PG_SIZ% = 1024

'ポジション・ブロック
Public ITEM_LOC_POS       As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type ITEM_LOCREC_Tag
    No(0 To 7)                  As Byte     'No
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
    
    IRI_QTY(0 To 7)             As Byte     '印刷入り数

    BIKOU(0 To 19)              As Byte     '印刷備考

    SOKO(0 To 1)                As Byte     '倉庫
    Retu(0 To 1)                As Byte     '列
    Ren(0 To 1)                 As Byte     '連
    Dan(0 To 1)                 As Byte     '段
    
    Print_SU(0 To 7)            As Byte     '印刷枚数

    FILLER(0 To 53)             As Byte
        
End Type
'データ・バッファ
Public ITEM_LOCREC              As ITEM_LOCREC_Tag

'キー定義

Type KEY0_ITEM_LOC                          'ＫＥＹ０
    No(0 To 7)                  As Byte     'No
End Type


Type KEY1_ITEM_LOC                          'ＫＥＹ１
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
End Type


Type KEY2_ITEM_LOC                          'ＫＥＹ２
    SOKO(0 To 1)                As Byte     '倉庫
    Retu(0 To 1)                As Byte     '列
    Ren(0 To 1)                 As Byte     '連
    Dan(0 To 1)                 As Byte     '段
    JGYOBU(0 To 0)              As Byte     '事業部区分
    NAIGAI(0 To 0)              As Byte     '国内外
    HIN_GAI(0 To 19)            As Byte     '品番（外部）
End Type




'キー・データ
Public K0_ITEM_LOC              As KEY0_ITEM_LOC
Public K1_ITEM_LOC              As KEY1_ITEM_LOC
Public K2_ITEM_LOC              As KEY2_ITEM_LOC

Type ITEM_LOC_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck
    ks5     As BtKeySpeck
    ks6     As BtKeySpeck
    ks7     As BtKeySpeck
    ks8     As BtKeySpeck
    ks9     As BtKeySpeck
    ks10    As BtKeySpeck
End Type

Private ITEM_LOC_Speck  As ITEM_LOC_FSpeck

Private Function ITEM_LOC_Create() As Integer
'********************************************************************
'*
'*              品目－棚マスタ  ファイル作成
'*
'*          CREATE 2012.06.01
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    ITEM_LOC_Create = True
                                            '品目－棚マスタフルパス取込み
    sts = GetIni(App.EXEName, ITEM_LOC_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, App.EXEName & " " & ITEM_LOC_ID & "読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    ITEM_LOC_Speck.fs.recoleng = Len(ITEM_LOCREC)       ' レコード長
    ITEM_LOC_Speck.fs.PageSize = ITEM_LOC_PG_SIZ        ' ページサイズ
    ITEM_LOC_Speck.fs.idexnumb = 3                      ' インデックス数
    ITEM_LOC_Speck.fs.fileflag = 0                      ' ファイルフラグ
    ITEM_LOC_Speck.fs.reserve = &H0                     ' 予約済み
'-----------------------------------------------
                                                ' キー０
    ITEM_LOC_Speck.ks0.keypos = 1                       ' キーポジション
    ITEM_LOC_Speck.ks0.keyleng = 8                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks0.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks0.reserve = &H0                    ' 予約済み
'-----------------------------------------------


'-----------------------------------------------
                                                ' キー１
    ITEM_LOC_Speck.ks1.keypos = 9                       ' キーポジション
    ITEM_LOC_Speck.ks1.keyleng = 1                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks1.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks1.reserve = &H0                    ' 予約済み

    ITEM_LOC_Speck.ks2.keypos = 10                      ' キーポジション
    ITEM_LOC_Speck.ks2.keyleng = 1                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks2.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks2.reserve = &H0                    ' 予約済み

    ITEM_LOC_Speck.ks3.keypos = 11                      ' キーポジション
    ITEM_LOC_Speck.ks3.keyleng = 20                     ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks3.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks3.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks3.reserve = &H0                    ' 予約済み
'-----------------------------------------------

'-----------------------------------------------
                                                ' キー２
    ITEM_LOC_Speck.ks4.keypos = 59                      ' キーポジション
    ITEM_LOC_Speck.ks4.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks4.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks4.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks4.reserve = &H0                    ' 予約済み
    
    ITEM_LOC_Speck.ks5.keypos = 61                      ' キーポジション
    ITEM_LOC_Speck.ks5.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks5.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks5.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks5.reserve = &H0                    ' 予約済み
    
    ITEM_LOC_Speck.ks6.keypos = 63                      ' キーポジション
    ITEM_LOC_Speck.ks6.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks6.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks6.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks6.reserve = &H0                    ' 予約済み
    
    ITEM_LOC_Speck.ks7.keypos = 65                      ' キーポジション
    ITEM_LOC_Speck.ks7.keyleng = 2                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks7.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks7.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks7.reserve = &H0                    ' 予約済み
    
    
    
    ITEM_LOC_Speck.ks8.keypos = 9                       ' キーポジション
    ITEM_LOC_Speck.ks8.keyleng = 1                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks8.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks8.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks8.reserve = &H0                    ' 予約済み

    ITEM_LOC_Speck.ks9.keypos = 10                      ' キーポジション
    ITEM_LOC_Speck.ks9.keyleng = 1                      ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks9.keyflag = BtKfExt + BtKfChg + BtKfDup + BtKfSeg
    ITEM_LOC_Speck.ks9.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks9.reserve = &H0                    ' 予約済み

    ITEM_LOC_Speck.ks10.keypos = 11                      ' キーポジション
    ITEM_LOC_Speck.ks10.keyleng = 20                     ' キー長
                                                        ' キーフラグ
    ITEM_LOC_Speck.ks10.keyflag = BtKfExt + BtKfChg + BtKfDup
    ITEM_LOC_Speck.ks10.keytype = Chr(BtKtString)        ' キータイプ
    ITEM_LOC_Speck.ks10.reserve = &H0                    ' 予約済み
'-----------------------------------------------



    sts = BTRV(BtOpCreate, ITEM_LOC_POS, ITEM_LOC_Speck, Len(ITEM_LOC_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "原産マスタ")
        Exit Function
    End If

    ITEM_LOC_Create = False

End Function

Public Function ITEM_LOC_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品目－棚マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    ITEM_LOC_Open = True
                                            '原産マスタフルパス取込み
    sts = GetIni(App.EXEName, ITEM_LOC_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, App.EXEName & " " & ITEM_LOC_ID & "読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = ITEM_LOC_Create()        '原産マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "原産マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "原産マスタ")
                Exit Function
        End Select
    Loop

    ITEM_LOC_Open = False

End Function

