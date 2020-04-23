Attribute VB_Name = "P_KANRI"
Option Explicit
'********************************************************************
'*                                                                  *
'*              管理マスタ  ファイル定義                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_KANRI_ID$ = "P_KANRI"

'ページサイズ
Private Const P_KANRI_PG_SIZ% = 512

'ポジション・ブロック
Public P_KANRI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_KANRIREC_Tag
    REC_NO(0 To 1)          As Byte         'ﾚｺｰﾄﾞ№
    SHIME_DD(0 To 1)        As Byte         'SDC月末締め日
    xSASHIZU_NO(0 To 4)     As Byte         '指図票№(現在値+1) 未使用とする 2007.11.28
    ORDER_NO(0 To 4)        As Byte         '発注№(現在値+1)
    URIAGE_NO(0 To 4)       As Byte         '資材売上ﾚｺｰﾄﾞ№(現在値+1)
    
    ZEI_CHANGE_YMD(0 To 7)  As Byte         '消費税変更日付
    NOW_ZEI_RITU(0 To 3)    As Byte         '現　消費税率
    NOW_MARUME(0 To 0)      As Byte         '    丸め
    NEW_ZEI_RITU(0 To 3)    As Byte         '新　消費税率
    NEW_MARUME(0 To 0)      As Byte         '    丸め
    
    SHONIN_CODE(0 To 4)     As Byte         '承認者ｺｰﾄﾞ
    KAISHA_NAME(0 To 29)    As Byte         '会社名
    CENTER_NAME(0 To 29)    As Byte         'センター名
    TEL_NO(0 To 14)         As Byte         '電話番号
    FAX_NO(0 To 14)         As Byte         'FAX番号
    
    URI_MARUME(0 To 0)      As Byte         '売上金額丸め
    SHI_MARUME(0 To 0)      As Byte         '仕入金額丸め
    
    SASHIZU_NO(0 To 7)      As Byte         '発注№(現在値+1)   2007.11.28
    
    
    NYUKO_S_RATE(0 To 6)    As Byte         '入庫　分レート     2008.02.13
    NYUKO_R_RATE(0 To 6)    As Byte         '入庫　余裕率       2008.02.13
    
    SYUKO_S_RATE(0 To 6)    As Byte         '出庫　分レート     2008.02.13
    SYUKO_R_RATE(0 To 6)    As Byte         '出庫　余裕率       2008.02.13
    
    SYUKA_S_RATE(0 To 6)    As Byte         '出庫　分レート     2008.02.13
    SYUKA_R_RATE(0 To 6)    As Byte         '出庫　余裕率       2008.02.13
    
    KOUTEI_LOT(0 To 5)      As Byte         '工程　前後工程標準ロット   2008.02.13
    KOUTEI_S_RATE(0 To 6)   As Byte         '工程　分レート             2008.02.13
    KOUTEI_R_RATE(0 To 6)   As Byte         '工程　余裕率               2008.02.13
    KOUTEI_SHIZAI(0 To 2)   As Byte         '工程　副資材確認点数       2008.02.13
    KOUTEI_BUHIN(0 To 2)    As Byte         '工程　同梱部品確認点数     2008.02.13
    KOUTEI_LABEL(0 To 2)    As Byte         '工程　ラベル貼付枚数       2008.02.13
    
    MITSUMORI_NO(0 To 7)    As Byte         '見積書№   2008.02.13
    SEIKYU_NO(0 To 7)       As Byte         '請求書№   2008.02.13
        
    
    MIN_URIAGE_NO(0 To 7)   As Byte         'ミニマム売上№     2008.02.13
    
    
    FILLER(0 To 18)         As Byte         'FILLER
End Type
'データ・バッファ
Public P_KANRIREC           As P_KANRIREC_Tag




Private Type P_KOTEI_Tag                    '2008.02.13
    KOTEI(0 To 2)       As Byte
End Type

Public Type P_KANRIREC02_Tag                '2008.02.13
    REC_NO(0 To 1)          As Byte         'ﾚｺｰﾄﾞ№
        
    BEF_KOTEI(0 To 9)       As P_KOTEI_Tag  '前工程
    MAIN_KOTEI(0 To 9)      As P_KOTEI_Tag  '作業工程
    AFT_KOTEI(0 To 9)       As P_KOTEI_Tag  '後工程
        
    FUTAI_KOTEI(0 To 4)     As P_KOTEI_Tag  '付帯工程　(現在未使用)
    KEIHI(0 To 4)           As P_KOTEI_Tag  '経費　(現在未使用)
    
    FILLER(0 To 133)        As Byte         'FILLER
End Type
'データ・バッファ
Public P_KANRIREC02         As P_KANRIREC02_Tag



'キー定義

Type KEY0_P_KANRI           'ＫＥＹ０
    REC_NO(0 To 1)          As Byte         'ﾚｺｰﾄﾞ№
End Type
    
'キー・データ
Public K0_P_KANRI           As KEY0_P_KANRI

Type P_KANRI_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_KANRI_Speck       As P_KANRI_FSpeck
Private Function P_KANRI_Create() As Integer
'********************************************************************
'*                                                                  *
'*              管理マスタ  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_KANRI_Create = True
                                            '管理マスタフルパス取込み
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_KANRI_Speck.fs.recoleng = Len(P_KANRIREC)            ' レコード長
    P_KANRI_Speck.fs.PageSize = P_KANRI_PG_SIZ          ' ページサイズ
    P_KANRI_Speck.fs.idexnumb = 1                       ' インデックス数
    P_KANRI_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_KANRI_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_KANRI_Speck.ks0.keypos = 1                        ' キーポジション
    P_KANRI_Speck.ks0.keyleng = 2                       ' キー長
    P_KANRI_Speck.ks0.keyflag = BtKfExt                 ' キーフラグ
    P_KANRI_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_KANRI_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    sts = BTRV(BtOpCreate, P_KANRI_POS, P_KANRI_Speck, Len(P_KANRI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "管理マスタ")
        Exit Function
    End If
    
    P_KANRI_Create = False

End Function

Public Function P_KANRI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              管理マスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_KANRI_Open = True
                                            '管理マスタフルパス取込み
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_KANRI_Create()      '管理マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "管理マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "管理マスタ")
                Exit Function
        End Select
    Loop
    
    P_KANRI_Open = False

End Function
Public Function P_KANRI_MAKE_Proc() As Integer
'----------------------------------------------------------------------------
'                   管理マスタの自動作成
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    P_KANRI_MAKE_Proc = True

    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)     'ﾚｺｰﾄﾞ№
    Call UniCode_Conv(P_KANRIREC.SHIME_DD, "31")            '月末締め日
    Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000")       '指図票№
    Call UniCode_Conv(P_KANRIREC.ORDER_NO, "00000")         '発注№
    Call UniCode_Conv(P_KANRIREC.URIAGE_NO, "00000")        '資材売上ﾚｺｰﾄﾞ№

    Call UniCode_Conv(P_KANRIREC.ZEI_CHANGE_YMD, "")        '消費税変更日付
    Call UniCode_Conv(P_KANRIREC.NOW_ZEI_RITU, "00.0")      '現　消費税率
    Call UniCode_Conv(P_KANRIREC.NOW_MARUME, "0")           '現　丸め
    Call UniCode_Conv(P_KANRIREC.NEW_ZEI_RITU, "00.0")      '新　消費税率
    Call UniCode_Conv(P_KANRIREC.NEW_MARUME, "0")           '新　丸め

    Call UniCode_Conv(P_KANRIREC.SHONIN_CODE, "")           '承認者ｺｰﾄﾞ
    Call UniCode_Conv(P_KANRIREC.KAISHA_NAME, "")           '会社名称
    Call UniCode_Conv(P_KANRIREC.TEL_NO, "")                '電話番号
    Call UniCode_Conv(P_KANRIREC.FAX_NO, "")                'FAX番号
    
    Call UniCode_Conv(P_KANRIREC.FILLER, "")

    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(BtOpInsert, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "管理マスタ")
                Exit Function
        End Select
    Loop
    
    
    P_KANRI_MAKE_Proc = False



End Function

