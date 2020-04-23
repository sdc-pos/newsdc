Attribute VB_Name = "PN_M"
Option Explicit
'********************************************************************
'*
'*              PNマスタ ファイル定義
'*
'*          CREATE 2009.05.29
'********************************************************************
'ファイルＩＤ
Public Const PN_M_ID = "PN_M"

'ページサイズ
Public Const PN_M_PG_SIZ% = 4096

'ポジション・ブロック
Public PN_M_POS As POSBLK
'=
'=
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type PN_MREC_Tag
    JCode(0 To 7)           As Byte     '事業場コード
    ShisanJCode(0 To 7)     As Byte     '資産管理事業場コード
    PN(0 To 19)             As Byte     '品目番号
    DModel(0 To 2)          As Byte     '代表機種品目コード
    HINMOKU(0 To 2)         As Byte     '品目コード
    SOKO(0 To 1)            As Byte     '倉庫コード
    KKeitai_10(0 To 9)      As Byte     '個装形態コード         2012.03.06 10桁
    Size_Kbn(0 To 0)        As Byte     '部品サイズ区分
    Saisu(0 To 13)          As Byte     '部品梱包才数
    TekiLabel(0 To 0)       As Byte     '適用機種ラベル発行区分
    KobaiTanto_5(0 To 4)      As Byte     '購買担当者コード     2012.03.06　5桁
    UnitKbn(0 To 0)         As Byte     'ユニット部品区分
    NaiKbn(0 To 0)          As Byte     '国内供給部品区分
    GaiKbn(0 To 0)          As Byte     '海外供給部品区分
    PnBetsu(0 To 39)        As Byte     '品目別名                                       2009.07.14 Byte数拡張 20 → 40
    PName(0 To 39)          As Byte     '品目名                                         2009.07.14 Byte数拡張 20 → 40
    Tanka2(0 To 9)          As Byte     '売上単価２
    Tanka3(0 To 9)          As Byte     '売上単価３
    Tanka4(0 To 9)          As Byte     '売上単価４
    Loc1(0 To 9)            As Byte     'ロケーション番号１
    Loc2(0 To 9)            As Byte     'ロケーション番号２
    Loc3(0 To 9)            As Byte     'ロケーション番号３
    SPn(0 To 19)            As Byte     '工場品目番号
    MadeIn(0 To 19)          As Byte     '現物表示原産国                        2009.07.14 Byte数拡張 10 → 20
    HyoTan(0 To 9)          As Byte     '標準単価
    Syutan(0 To 1)                      As Byte     '終端文字
        INS_ID(0 To 9)                  As Byte     '登録ID                                             2009.07.14 追加
        INS_TM(0 To 11)                 As Byte     '登録日時 yyyymmddhhmm              2009.07.14 追加
        UPD_ID(0 To 9)                  As Byte     '更新ID                                             2009.07.14 追加
        UPD_TM(0 To 11)                 As Byte     '更新日時 yyyymmddhhmm              2009.07.14 追加
    
    
    MadeInCode(0 To 2)      As Byte                 '2010.08.20
    GENSANKOKU(0 To 19)     As Byte                 '原産国　2012.02.06
        
    
    KKeitai(0 To 13)         As Byte     '個装形態               2012.03.07
    KobaiTanto(0 To 7)      As Byte     '購買担当者ｺｰﾄﾞ         2012.03.07
    NaiModel(0 To 19)       As Byte     '国内機種品目番号       2012.03.07
    NaiModelNew(0 To 19)    As Byte     '国内最新機種品目番号   2012.03.07
    GaiModel(0 To 19)       As Byte     '輸出機種品目番号       2012.03.07
    GaiModelNew(0 To 19)    As Byte     '輸出最新機種品目番号   2012.03.07
    PNameEngA(0 To 39)      As Byte     '英語 品目別名          2012.03.07
    PNameEng(0 To 39)       As Byte     '英語 品目名            2012.03.07
    NaiDisconYm(0 To 5)     As Byte     '国内供給打切年月       2012.03.07
    GaiDisconYm(0 To 5)     As Byte     '海外供給打切年月       2012.03.07
    
    
    
    
    
    
    
    'FILLER(0 To 10)         As Byte
    

End Type
'データ・バッファ
Public PN_MREC           As PN_MREC_Tag


'キー定義
Type KEY0_PN_M                       'ＫＥＹ０
    JCode(0 To 7)           As Byte     '事業場コード
    ShisanJCode(0 To 7)     As Byte     '資産管理事業場コード
    PN(0 To 19)             As Byte     '品目番号
End Type

Type KEY1_PN_M                       'ＫＥＹ１
    JCode(0 To 7)           As Byte     '事業場コード
    ShisanJCode(0 To 7)     As Byte     '資産管理事業場コード
    SPn(0 To 19)            As Byte     '工場品目番号
End Type


'キー・データ
Public K0_PN_M           As KEY0_PN_M
Public K1_PN_M           As KEY1_PN_M

Private Type PN_M_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
    ks2 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private PN_M_Speck    As PN_M_FSpeck
Private Function PN_M_Create() As Integer
'********************************************************************
'*
'*              PN_M管理集計ファイル  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.04.22
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PN_M_Create = True
                                            'PN_M管理集計ファイルフルパス取込み
    sts = GetIni("FILE", PN_M_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[PN_M] 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PN_M_Speck.fs.recoleng = Len(PN_MREC)             ' レコード長
    PN_M_Speck.fs.PageSize = PN_M_PG_SIZ              ' ページサイズ
    PN_M_Speck.fs.idexnumb = 2                       ' インデックス数
    PN_M_Speck.fs.fileflag = 0                       ' ファイルフラグ
    PN_M_Speck.fs.reserve = &H0                      ' 予約済み

'---------------------------------------------------' キー０
    PN_M_Speck.ks0.keypos = 1                        ' キーポジション
    PN_M_Speck.ks0.keyleng = 8 + 8 + 20              ' キー長
                                                     ' キーフラグ
    PN_M_Speck.ks0.keyflag = BtKfExt + BtKfChg
    PN_M_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    PN_M_Speck.ks0.reserve = &H0                     ' 予約済み

'---------------------------------------------------' キー１
    PN_M_Speck.ks1.keypos = 1                        ' キーポジション
    PN_M_Speck.ks1.keyleng = 8 + 8                     ' キー長
    PN_M_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' キーフラグ
    PN_M_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    PN_M_Speck.ks1.reserve = &H0                     ' 予約済み

    PN_M_Speck.ks2.keypos = 219                      ' キーポジション
    PN_M_Speck.ks2.keyleng = 20                      ' キー長
    PN_M_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg                ' キーフラグ
    PN_M_Speck.ks2.keytype = Chr(BtKtString)         ' キータイプ
    PN_M_Speck.ks2.reserve = &H0                     ' 予約済み
    
    
    sts = BTRV(BtOpCreate, PN_M_POS, PN_M_Speck, Len(PN_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "PNマスタ")
        Exit Function
    End If

    PN_M_Create = False

End Function

Function PN_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              PNマスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.04.22
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PN_M_Open = True
                                            'PN_M管理集計ファイルフルパス取込み
    sts = GetIni("FILE", PN_M_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PN_M_POS, PN_MREC, Len(PN_MREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PN_M_Create()        'PN_M管理集計ファイル作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PN_M_POS, PN_MREC, Len(PN_MREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "PNマスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "PNマスタ")
                Exit Function
        End Select
    Loop
    PN_M_Open = False
End Function

Public Function PN_M_GET(JG As String, PN As String, Locked As Integer) As Integer
'           引数
'   JG      事業部
'   PN      品目番号

'   Locked  ＧｅｔＬｏｃｋ

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim JC          As String       '事業場コード
Dim SC          As String       '資産管理事業場コード


    PN_M_GET = True
    
    '           事業場コード　設定
    JC = String(UBound(PN_MREC.JCode) + 1, "0")
    If GetIni("JCODE", JG, "PN_JCode", JC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] 事業部[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(JC) = "" Then
        JC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    '           資産管理事業場コード　設定
    SC = String(UBound(PN_MREC.ShisanJCode) + 1, "0")
    If GetIni("ShisanJCode", JG, "PN_JCode", SC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] 事業部[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(SC) = "" Then
        SC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    
    Call UniCode_Conv(K0_PN_M.JCode, JC)        '事業場コード
    Call UniCode_Conv(K0_PN_M.ShisanJCode, SC)  '資産管理事業場コード
    Call UniCode_Conv(K0_PN_M.PN, PN)           '品目番号
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'MsgBox "指定されたデータがありません。"
                
                
                Call UniCode_Conv(PN_MREC.PN, PN)
                Call UniCode_Conv(PN_MREC.SPn, PN)
                
                                  
                
                Call UniCode_Conv(PN_MREC.PName, "")                '品目名
                Call UniCode_Conv(PN_MREC.Tanka2, "0000000.00")     '売上単価２
                Call UniCode_Conv(PN_MREC.Tanka3, "0000000.00")     '売上単価３
                Call UniCode_Conv(PN_MREC.Tanka4, "0000000.00")     '売上単価４

                Call UniCode_Conv(PN_MREC.SPn, "")                  '工場品目番号
                
                
                
                PN_M_GET = BtErrKeyNotFound
                
                
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<PN_M>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "PN_M")
                Exit Function
        End Select
    Loop
    
    PN_M_GET = False

End Function

Public Function PN_M_GET2(JG As String, PN As String, Locked As Integer) As Integer
'           引数
'   JG      事業部
'   PN      品目番号

'   Locked  ＧｅｔＬｏｃｋ

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim JC          As String       '事業場コード
Dim SC          As String       '資産管理事業場コード


    PN_M_GET2 = True
    
    If Trim(PN) = "" Then
        'MsgBox "品目番号＝空白　→　指定されたデータがありません。"
        Exit Function
    End If
    
    
    '           事業場コード　設定
    JC = String(UBound(PN_MREC.JCode) + 1, "0")
    If GetIni("JCODE", JG, "PN_JCode", JC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] 事業部[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(JC) = "" Then
        JC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    '           資産管理事業場コード　設定
    SC = String(UBound(PN_MREC.ShisanJCode) + 1, "0")
    If GetIni("ShisanJCode", JG, "PN_JCode", SC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] 事業部[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(SC) = "" Then
        SC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    
    Call UniCode_Conv(K1_PN_M.JCode, JC)        '事業場コード
    Call UniCode_Conv(K1_PN_M.ShisanJCode, SC)  '資産管理事業場コード
    Call UniCode_Conv(K1_PN_M.SPn, PN)          '工場品目番号
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, PN_M_POS, PN_MREC, Len(PN_MREC), K1_PN_M, Len(K1_PN_M), 1)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'MsgBox "指定されたデータがありません。"
                
                
                Call UniCode_Conv(PN_MREC.PN, PN)
                Call UniCode_Conv(PN_MREC.SPn, PN)
                
                Call UniCode_Conv(PN_MREC.PName, "")                '品目名
                Call UniCode_Conv(PN_MREC.Tanka2, "0000000.00")     '売上単価２
                Call UniCode_Conv(PN_MREC.Tanka3, "0000000.00")     '売上単価３
                Call UniCode_Conv(PN_MREC.Tanka4, "0000000.00")     '売上単価４
                
                
                
                PN_M_GET2 = BtErrKeyNotFound

                
                
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<PN_M>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "PN_M")
                Exit Function
        End Select
    Loop
    
    PN_M_GET2 = False

End Function


Public Sub Rclr_PN_MREC()
'********************************************************************
'*
'*              PNマスタ  レコード初期化
'*
'********************************************************************

    Call UniCode_Conv(PN_MREC.JCode, "")                '事業場コード
    Call UniCode_Conv(PN_MREC.ShisanJCode, "")          '資産管理事業場コード
    Call UniCode_Conv(PN_MREC.PN, "")                   '品目番号
    Call UniCode_Conv(PN_MREC.DModel, "")               '代表機種品目コード
    Call UniCode_Conv(PN_MREC.HINMOKU, "")              '品目コード
    Call UniCode_Conv(PN_MREC.SOKO, "")                 '倉庫コード
    Call UniCode_Conv(PN_MREC.KKeitai, "")              '個装形態コード
    Call UniCode_Conv(PN_MREC.Size_Kbn, "")             '部品サイズ区分
    Call UniCode_Conv(PN_MREC.Saisu, "")                '部品梱包才数
    Call UniCode_Conv(PN_MREC.TekiLabel, "")            '適用機種ラベル発行区分
    
    Call UniCode_Conv(PN_MREC.KobaiTanto, "")           '購買担当者コード
    Call UniCode_Conv(PN_MREC.UnitKbn, "")              'ユニット部品区分
    Call UniCode_Conv(PN_MREC.NaiKbn, "")               '国内供給部品区分
    Call UniCode_Conv(PN_MREC.GaiKbn, "")               '海外供給部品区分
    Call UniCode_Conv(PN_MREC.PnBetsu, "")              '品目別名
    Call UniCode_Conv(PN_MREC.PName, "")                '品目名
    Call UniCode_Conv(PN_MREC.Tanka2, "")                '売上単価２
    Call UniCode_Conv(PN_MREC.Tanka3, "")                '売上単価３
    Call UniCode_Conv(PN_MREC.Tanka4, "")                '売上単価４
    Call UniCode_Conv(PN_MREC.Loc1, "")                'ロケーション番号１
    
    Call UniCode_Conv(PN_MREC.Loc2, "")                'ロケーション番号２
    Call UniCode_Conv(PN_MREC.Loc3, "")                'ロケーション番号３
    Call UniCode_Conv(PN_MREC.SPn, "")                '工場品目番号
    Call UniCode_Conv(PN_MREC.MadeIn, "")                '現物表示原産国
    Call UniCode_Conv(PN_MREC.HyoTan, "")                '標準単価
    Call UniCode_Conv(PN_MREC.Syutan, "")                '終端文字
    'Call UniCode_Conv(PN_MREC.FILLER, "")                '


End Sub

