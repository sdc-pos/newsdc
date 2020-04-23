Attribute VB_Name = "FURIKAE"
Option Explicit
'********************************************************************
'*
'*              品番振替Ｍ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const FURIKAE_ID$ = "FURIKAE"

'ページサイズ
Public Const FURIKAE_PG_SIZ% = 1024

'ポジション・ブロック
Public FURIKAE_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type FURIKAEREC_Tag
    JGYOBU_MAE(0 To 0)          As Byte     '振替前事業部           2012.03.13
    NAIGAI_MAE(0 To 0)          As Byte     '振替前国内外           2012.03.13
    HIN_MAE(0 To 19)            As Byte     '振替前品番（外部）
    JGYOBU_GO(0 To 0)           As Byte     '振替後事業部           2012.03.13
    NAIGAI_GO(0 To 0)           As Byte     '振替後国内外           2012.03.13
    HIN_GO(0 To 19)             As Byte     '振替後品番（外部）
    BIKOU(0 To 39)              As Byte     '備考
    
    CUT_SU(0 To 2)              As Byte     '切断数                 2012.03.14
    
    
    MOTO_LEN(0 To 2)            As Byte     '元の長さ               2012.12.26
    
    
    KO_QTY(0 To 3)              As Byte     '員数                   2013.02.22
    
    
    FILLER(0 To 17)             As Byte    '                        2013.02.22 桁数変更
    
    INS_TANTO(0 To 9)           As Byte     '追加　担当者
    Ins_DateTime(0 To 13)       As Byte     '追加　日時

    UPD_TANTO(0 To 9)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時

End Type

'データ・バッファ
Public FURIKAEREC   As FURIKAEREC_Tag

'キー定義
Type KEY0_FURIKAE            'ＫＥＹ０
    JGYOBU_MAE(0 To 0)                  As Byte     '振替前事業部           2012.03.13
    NAIGAI_MAE(0 To 0)                  As Byte     '振替前国内外           2012.03.13
    HIN_MAE(0 To 19)                    As Byte     '振替前品番（外部）
    JGYOBU_GO(0 To 0)                   As Byte     '振替後事業部           2012.03.13
    NAIGAI_GO(0 To 0)                   As Byte     '振替後国内外           2012.03.13
    HIN_GO(0 To 19)                     As Byte     '振替後品番（外部）
End Type

Type KEY1_FURIKAE            'ＫＥＹ１
    JGYOBU_GO(0 To 0)                   As Byte     '振替後事業部           2012.03.13
    NAIGAI_GO(0 To 0)                   As Byte     '振替後国内外           2012.03.13
    HIN_GO(0 To 19)                     As Byte     '振替後品番（外部）
    JGYOBU_MAE(0 To 0)                  As Byte     '振替前事業部           2012.03.13
    NAIGAI_MAE(0 To 0)                  As Byte     '振替前国内外           2012.03.13
    HIN_MAE(0 To 19)                    As Byte     '振替前品番（外部）
End Type


'キー・データ
Public K0_FURIKAE                   As KEY0_FURIKAE
Public K1_FURIKAE                   As KEY1_FURIKAE

Type FURIKAE_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
    ks3     As BtKeySpeck
    ks4     As BtKeySpeck           '2012.03.13
    ks5     As BtKeySpeck           '2012.03.13
    ks6     As BtKeySpeck           '2012.03.13
    ks7     As BtKeySpeck           '2012.03.13
    ks8     As BtKeySpeck           '2012.03.13
    ks9     As BtKeySpeck           '2012.03.13
    ks10    As BtKeySpeck           '2012.03.13
    ks11    As BtKeySpeck           '2012.03.13

End Type

Private FURIKAE_Speck               As FURIKAE_FSpeck
Private Function FURIKAE_Create() As Integer
'********************************************************************
'*
'*              品番振替Ｍ　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    FURIKAE_Create = True
                                            '品番振替Ｍフルパス取込み
    sts = GetIni("FILE", FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [" & FURIKAE_ID & "]読み込みエラー")
        Exit Function
    End If
     
    FullPath = RTrim(c)
    
    FURIKAE_Speck.fs.recoleng = Len(FURIKAEREC)         ' レコード長
    FURIKAE_Speck.fs.PageSize = FURIKAE_PG_SIZ          ' ページサイズ
    FURIKAE_Speck.fs.idexnumb = 2                   ' インデックス数
    FURIKAE_Speck.fs.fileflag = 0                   ' ファイルフラグ
    FURIKAE_Speck.fs.reserve = &H0                  ' 予約済み
'-----------------------------------------------
                                                ' キー０
    FURIKAE_Speck.ks0.keypos = 1                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks0.keyleng = 1
                                                ' キーフラグ
    FURIKAE_Speck.ks0.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks0.reserve = &H0                 ' 予約済み


    FURIKAE_Speck.ks1.keypos = 2                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks1.keyleng = 1
                                                ' キーフラグ
    FURIKAE_Speck.ks1.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks1.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks1.reserve = &H0                 ' 予約済み

    FURIKAE_Speck.ks2.keypos = 3                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks2.keyleng = 20
                                                ' キーフラグ
    FURIKAE_Speck.ks2.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks2.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks2.reserve = &H0                 ' 予約済み

    FURIKAE_Speck.ks3.keypos = 23                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks3.keyleng = 1
                                                ' キーフラグ
    FURIKAE_Speck.ks3.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks3.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks3.reserve = &H0                 ' 予約済み


    FURIKAE_Speck.ks4.keypos = 24                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks4.keyleng = 1
                                                ' キーフラグ
    FURIKAE_Speck.ks4.keyflag = BtKfExt + BtKfSeg '+ BtKfDup
    FURIKAE_Speck.ks4.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks4.reserve = &H0                 ' 予約済み


    FURIKAE_Speck.ks5.keypos = 25                ' キーポジション
                                                ' キー長
    FURIKAE_Speck.ks5.keyleng = 20
                                                ' キーフラグ
    FURIKAE_Speck.ks5.keyflag = BtKfExt  '+ BtKfDup
    FURIKAE_Speck.ks5.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks5.reserve = &H0                 ' 予約済み

'-----------------------------------------------
                                                ' キー１
    FURIKAE_Speck.ks6.keypos = 23                   ' キーポジション
    FURIKAE_Speck.ks6.keyleng = 1                   ' キー長
                                                ' キーフラグ
    FURIKAE_Speck.ks6.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks6.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks6.reserve = &H0                 ' 予約済み

    FURIKAE_Speck.ks7.keypos = 24                   ' キーポジション
    FURIKAE_Speck.ks7.keyleng = 1                   ' キー長
                                                ' キーフラグ
    FURIKAE_Speck.ks7.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks7.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks7.reserve = &H0                 ' 予約済み

    FURIKAE_Speck.ks8.keypos = 25                   ' キーポジション
    FURIKAE_Speck.ks8.keyleng = 20                   ' キー長
                                                ' キーフラグ
    FURIKAE_Speck.ks8.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks8.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks8.reserve = &H0                 ' 予約済み



    FURIKAE_Speck.ks9.keypos = 1                   ' キーポジション
    FURIKAE_Speck.ks9.keyleng = 1                   ' キー長
                                                ' キーフラグ
    FURIKAE_Speck.ks9.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks9.keytype = Chr(BtKtString)     ' キータイプ
    FURIKAE_Speck.ks9.reserve = &H0                 ' 予約済み

    FURIKAE_Speck.ks10.keypos = 2                   ' キーポジション
    FURIKAE_Speck.ks10.keyleng = 1                  ' キー長
                                                    ' キーフラグ
    FURIKAE_Speck.ks10.keyflag = BtKfExt + BtKfSeg
    FURIKAE_Speck.ks10.keytype = Chr(BtKtString)    ' キータイプ
    FURIKAE_Speck.ks10.reserve = &H0                ' 予約済み

    FURIKAE_Speck.ks11.keypos = 3                   ' キーポジション
    FURIKAE_Speck.ks11.keyleng = 20                 ' キー長
                                                    ' キーフラグ
    FURIKAE_Speck.ks11.keyflag = BtKfExt
    FURIKAE_Speck.ks11.keytype = Chr(BtKtString)    ' キータイプ
    FURIKAE_Speck.ks11.reserve = &H0                ' 予約済み


'-----------------------------------------------

    sts = BTRV(BtOpCreate, FURIKAE_POS, FURIKAE_Speck, Len(FURIKAE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "品番振替Ｍ")
        Exit Function
    End If

    FURIKAE_Create = False

End Function

Public Function FURIKAE_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              品番振替Ｍ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    FURIKAE_Open = True
                                            '品番振替Ｍフルパス取込み
    sts = GetIni("FILE", FURIKAE_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [FURIKAE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = FURIKAE_Create()        '品番振替Ｍ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "品番振替Ｍ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
    FURIKAE_Open = False
End Function


Function FURIKAE_Get(JGYOBU As String, NAIGAI As String, HIN_MAE As String, HIN_GO As String, Locked As Integer)
'----------------------------------------------------------------------------
'                   品番振替ＭファイルＧｅｔ

'       Locked      :False=NormalGet,ﾛｯｸ時はBtrieveｵﾍﾟﾚｰｼｮﾝのﾛｯｸ定数
'----------------------------------------------------------------------------
Dim com As Integer
Dim sts As Integer
Dim yn As Integer

    FURIKAE_Get = True
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, JGYOBU)    '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, NAIGAI)    '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, HIN_MAE)
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_GO, JGYOBU)     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_GO, NAIGAI)     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, HIN_GO)
    com = BtOpGetEqual + Locked
Do
    sts = BTRV(com, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    
    Select Case sts
        Case BtNoErr
            Exit Do
        Case BtErrKeyNotFound       'レコード無し
            
            'MsgBox "指定された工程がありません。"
            Exit Function
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
            yn = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                        "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
            If yn = vbNo Then Exit Function
        Case Else
            Call File_Error(sts, com, "品番振替Ｍ")
            Exit Function
    End Select
Loop

    FURIKAE_Get = False

End Function
Sub FURIKAE_CLOSE()
Dim sts As Integer

    sts = BTRV(BtOpClose, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品番振替Ｍ")
        End If
    End If

End Sub
Sub FURIKAE_CLR()


    Call UniCode_Conv(FURIKAEREC.JGYOBU_MAE, "")        '2012.03.13
    Call UniCode_Conv(FURIKAEREC.NAIGAI_MAE, "")        '2012.03.13
    Call UniCode_Conv(FURIKAEREC.HIN_MAE, "")
    
    Call UniCode_Conv(FURIKAEREC.JGYOBU_GO, "")         '2012.03.13
    Call UniCode_Conv(FURIKAEREC.NAIGAI_GO, "")         '2012.03.13
    Call UniCode_Conv(FURIKAEREC.HIN_GO, "")            '2012.03.14
    Call UniCode_Conv(FURIKAEREC.BIKOU, "")
    
    Call UniCode_Conv(FURIKAEREC.CUT_SU, "")
    
    
    Call UniCode_Conv(FURIKAEREC.FILLER, "")
    
    Call UniCode_Conv(FURIKAEREC.INS_TANTO, "")
    Call UniCode_Conv(FURIKAEREC.Ins_DateTime, "")
    Call UniCode_Conv(FURIKAEREC.UPD_TANTO, "")
    Call UniCode_Conv(FURIKAEREC.UPD_DATETIME, "")
    
    'Call UniCode_Conv(FURIKAEREC.FILLER, String(UBound(FURIKAEREC.FILLER) + 1, "0"))

End Sub

