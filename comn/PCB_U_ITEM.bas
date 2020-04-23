Attribute VB_Name = "PCB_U_ITEM"
Option Explicit
'********************************************************************
'*
'*              PCB.U設変  ファイル定義
'*
'*          CREATE 2014.06.18
'********************************************************************
'ファイルＩＤ
Public Const PCB_U_ITEM_ID$ = "PCB_U_ITEM"

'ページサイズ
Public Const PCB_U_ITEM_PG_SIZ% = 4096

'ポジション・ブロック
Public PCB_U_ITEM_POS               As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type PCB_U_ITEM_REC_Tag
    JGYOBU(0 To 0)                  As Byte     '事業部区分
    NAIGAI(0 To 0)                  As Byte     '国内外
    HIN_GAI(0 To 19)                As Byte     '品番（外部）
    
    NAI_BUHIN(0 To 0)               As Byte     '供給区分
    KENSA_JIGU(0 To 0)              As Byte     '検査治具
    
    HITUYO_SU(0 To 4)               As Byte     '必要数　個
    HITUYO_TUKI(0 To 1)             As Byte     '必要数　月
    
    SETUHEN_LAST_DATE(0 To 7)       As Byte     '設計変更最終日
    SENDO_LAST_DATE(0 To 7)         As Byte     '鮮度管理最終日
    
    MODULE_KBN(0 To 0)              As Byte     'モジュール対象区分
    SETUHEN_KBN(0 To 0)             As Byte     '設計変更対象区分
    
    
    KANRI_NO(0 To 1)                As Byte     '管理№
    EX_DATE(0 To 7)                 As Byte     '日付
    SETUHEN_NO(0 To 4)              As Byte     '設変管理№
    
    BEF_HIN_GAI(0 To 19)            As Byte     '変更前　ｻｰﾋﾞｽ品番
    BEF_HIN_NAI(0 To 19)            As Byte     '変更前　工場品番
    AFT_HIN_GAI(0 To 19)            As Byte     '変更前　ｻｰﾋﾞｽ品番
    AFT_HIN_NAI(0 To 19)            As Byte     '変更前　工場品番
    
    HEN_BUHIN(0 To 39)              As Byte     '変更部品
    HEN_NAIYO(0 To 49)              As Byte     '変更内容
    HEN_BASHO(0 To 19)              As Byte     '交換場所
    
    SETUHEN_HOKAN(0 To 19)          As Byte     '設変原紙保管
        
    BIKOU(0 To 199)                 As Byte     '備考
    
    
    FILLER(0 To 245)                As Byte         'FILLER
    INS_TANTO(0 To 9)               As Byte         '追加　担当者
    Ins_DateTime(0 To 13)           As Byte         '追加　日時
    UPD_TANTO(0 To 9)               As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)           As Byte         '更新　日時



End Type
'データ・バッファ
Public PCB_U_ITEM_REC               As PCB_U_ITEM_REC_Tag

'キー定義
Type KEY0_PCB_U_ITEM                'ＫＥＹ０
    JGYOBU(0 To 0)                  As Byte     '事業部区分
    NAIGAI(0 To 0)                  As Byte     '国内外
    HIN_GAI(0 To 19)                As Byte     '品番（外部）
End Type








'キー・データ
Public K0_PCB_U_ITEM                As KEY0_PCB_U_ITEM


Private Type PCB_U_ITEM_FSpeck
    fs      As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck

End Type

Private PCB_U_ITEM_Speck            As PCB_U_ITEM_FSpeck

Private Function PCB_U_ITEM_Create() As Integer
'********************************************************************
'*
'*              PCB.U設変  ファイル定義
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PCB_U_ITEM_Create = True
                                            'PCB.U設変　フルパス取込み
    sts = GetIni("FILE", PCB_U_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PCB_U_ITEM]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)

    PCB_U_ITEM_Speck.fs.recoleng = Len(PCB_U_ITEM_REC)      ' レコード長
    PCB_U_ITEM_Speck.fs.PageSize = PCB_U_ITEM_PG_SIZ        ' ページサイズ
    PCB_U_ITEM_Speck.fs.idexnumb = 1                        ' インデックス数
    PCB_U_ITEM_Speck.fs.fileflag = 0                        ' ファイルフラグ
    PCB_U_ITEM_Speck.fs.reserve = &H0                       ' 予約済み
'-----------------------------------------------
                                                ' キー０
    PCB_U_ITEM_Speck.ks0.keypos = 1                         ' キーポジション
    PCB_U_ITEM_Speck.ks0.keyleng = 1                        ' キー長
    PCB_U_ITEM_Speck.ks0.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PCB_U_ITEM_Speck.ks0.keytype = Chr(BtKtString)          ' キータイプ
    PCB_U_ITEM_Speck.ks0.reserve = &H0                      ' 予約済み

    PCB_U_ITEM_Speck.ks1.keypos = 2                         ' キーポジション
    PCB_U_ITEM_Speck.ks1.keyleng = 1                        ' キー長
                                                            
    PCB_U_ITEM_Speck.ks1.keyflag = BtKfExt + BtKfSeg        ' キーフラグ
    PCB_U_ITEM_Speck.ks1.keytype = Chr(BtKtString)          ' キータイプ
    PCB_U_ITEM_Speck.ks1.reserve = &H0                      ' 予約済み

    PCB_U_ITEM_Speck.ks2.keypos = 3                         ' キーポジション
    PCB_U_ITEM_Speck.ks2.keyleng = 20                       ' キー長
                                                            
    PCB_U_ITEM_Speck.ks2.keyflag = BtKfExt                  ' キーフラグ
    PCB_U_ITEM_Speck.ks2.keytype = Chr(BtKtString)          ' キータイプ
    PCB_U_ITEM_Speck.ks2.reserve = &H0                      ' 予約済み




'-----------------------------------------------

    sts = BTRV(BtOpCreate, PCB_U_ITEM_POS, PCB_U_ITEM_Speck, Len(PCB_U_ITEM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "PCB.U設変")
        Exit Function
    End If

    PCB_U_ITEM_Create = False

End Function

Public Function PCB_U_ITEM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              PCB.U設変  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PCB_U_ITEM_Open = True
                                            'PCB.U設変 フルパス取込み
    sts = GetIni("FILE", PCB_U_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PCB_U_ITEM]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, PCB_U_ITEM_POS, PCB_U_ITEM_REC, Len(PCB_U_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PCB_U_ITEM_Create()        'PCB.U設変作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PCB_U_ITEM_POS, PCB_U_ITEM_REC, Len(PCB_U_ITEM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "邸PCB.U設変")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "PCB.U設変")
                Exit Function
        End Select
    Loop

    PCB_U_ITEM_Open = False

End Function

