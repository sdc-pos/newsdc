Attribute VB_Name = "SE_MIN_URIAGE"
Option Explicit
'********************************************************************
'*
'*              ミニマム売上実績  ファイル定義
'*
'*          CREATE 2008.02.28
'********************************************************************
'ファイルＩＤ
Public Const SE_MIN_URIAGE_ID$ = "SE_MIN_URIAGE"

'ページサイズ
Public Const SE_MIN_URIAGE_PG_SIZ% = 4096

'ポジション・ブロック
Public SE_MIN_URIAGE_POS         As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SE_MIN_URIAGEREC_Tag
    JITU_DATE(0 To 7)           As Byte     '実績日付
    DEN_NO(0 To 7)              As Byte     '伝票№
    GYO_NO(0 To 2)              As Byte     '行№
    KEIJYO_YM(0 To 5)           As Byte     '計上年月(YYYYMM)
    UKEHARAI_CODE(0 To 4)       As Byte     '受払先ｺｰﾄﾞ
    
    SE_KBN(0 To 1)              As Byte     '請求区分
    MANA_KBN(0 To 1)            As Byte     '経営項目
    POST_CODE(0 To 1)           As Byte     '部署
    SUB_ITEM(0 To 39)           As Byte     '請求項目（提出用）
    SDC_ITEM(0 To 39)           As Byte     '請求項目（ＳＤＣ用）
        
    SURYO(0 To 11)              As Byte     '数量   S9(8)V99
    TANKA(0 To 10)              As Byte     '数量   9(8)V99
        
    URI_KIN(0 To 8)             As Byte     '金額   S9(9)
    ZEI_KIN(0 To 8)             As Byte     '消費税 S9(9)
        
    TEKIYO(0 To 39)             As Byte     '摘要
        
        
    UPD_TANTO(0 To 4)           As Byte     '更新　担当者
    UPD_DATETIME(0 To 13)       As Byte     '更新　日時
    
    
    FILLER(0 To 103)            As Byte     'FILLER




End Type
'データ・バッファ
Public SE_MIN_URIAGEREC         As SE_MIN_URIAGEREC_Tag

'キー定義

Type KEY0_SE_MIN_URIAGE         'ＫＥＹ０
    JITU_DATE(0 To 7)           As Byte     '実績日付
    DEN_NO(0 To 7)              As Byte     '伝票№
    GYO_NO(0 To 2)              As Byte     '行№
End Type





'キー・データ
Public K0_SE_MIN_URIAGE         As KEY0_SE_MIN_URIAGE

Type SE_MIN_URIAGE_FSpeck
    fs      As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
End Type

Private SE_MIN_URIAGE_Speck     As SE_MIN_URIAGE_FSpeck
Private Function SE_MIN_URIAGE_Create() As Integer
'********************************************************************
'*
'*              ミニマム売上実績  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_MIN_URIAGE_Create = True
                                            'ミニマム売上実績フルパス取込み
    sts = GetIni("FILE", SE_MIN_URIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_MIN_URIAGE]読み込みエラー ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    SE_MIN_URIAGE_Speck.fs.recoleng = Len(SE_MIN_URIAGEREC)     ' レコード長
    SE_MIN_URIAGE_Speck.fs.PageSize = SE_MIN_URIAGE_PG_SIZ      ' ページサイズ
    SE_MIN_URIAGE_Speck.fs.idexnumb = 1                         ' インデックス数
    SE_MIN_URIAGE_Speck.fs.fileflag = 0                         ' ファイルフラグ
    SE_MIN_URIAGE_Speck.fs.reserve = &H0                        ' 予約済み

'-----------------------------------------------
                                                ' キー１
    SE_MIN_URIAGE_Speck.ks0.keypos = 1                  ' キーポジション
    SE_MIN_URIAGE_Speck.ks0.keyleng = 8                 ' キー長
    SE_MIN_URIAGE_Speck.ks0.keyflag = BtKfExt + _
                                        BtKfSeg
    SE_MIN_URIAGE_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SE_MIN_URIAGE_Speck.ks0.reserve = &H0               ' 予約済み
                                                
    SE_MIN_URIAGE_Speck.ks1.keypos = 9                  ' キーポジション
    SE_MIN_URIAGE_Speck.ks1.keyleng = 8                 ' キー長
    SE_MIN_URIAGE_Speck.ks1.keyflag = BtKfExt + _
                                        BtKfSeg
    SE_MIN_URIAGE_Speck.ks1.keytype = Chr(BtKtString)   ' キータイプ
    SE_MIN_URIAGE_Speck.ks1.reserve = &H0               ' 予約済み

                                                
    SE_MIN_URIAGE_Speck.ks2.keypos = 17                  ' キーポジション
    SE_MIN_URIAGE_Speck.ks2.keyleng = 3                 ' キー長
    SE_MIN_URIAGE_Speck.ks2.keyflag = BtKfExt
    SE_MIN_URIAGE_Speck.ks2.keytype = Chr(BtKtString)   ' キータイプ
    SE_MIN_URIAGE_Speck.ks2.reserve = &H0               ' 予約済み


'-----------------------------------------------



    sts = BTRV(BtOpCreate, SE_MIN_URIAGE_POS, SE_MIN_URIAGE_Speck, Len(SE_MIN_URIAGE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ミニマム売上実績")
        Exit Function
    End If

    SE_MIN_URIAGE_Create = False

End Function

Public Function SE_MIN_URIAGE_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              ミニマム売上実績  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SE_MIN_URIAGE_Open = True
                                            'ミニマム売上実績フルパス取込み
    sts = GetIni("FILE", SE_MIN_URIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_MIN_URIAGE]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_MIN_URIAGE_Create()    'ミニマム売上実績作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ミニマム売上実績")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ミニマム売上実績")
                Exit Function
        End Select
    Loop

    SE_MIN_URIAGE_Open = False
    

End Function


