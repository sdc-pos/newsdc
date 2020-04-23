Attribute VB_Name = "ODR_TEMP3"
Option Explicit
'********************************************************************
'*                                                                  *
'*              中間　所要量Ｆ（WORK) ファイル定義              　　　*
'*                                                                  *
'*          CREATE 2008.03.06                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_TEMP3_ID$ = "ODR_TEMP3"

'ページサイズ
Private Const ODR_TEMP3_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_TP3_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_TP3_R_Tag
    USE_YM(0 To 5)              As Byte         '使用月
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    USE_QTY(0 To 10)            As Byte         '使用数         9(8)v9(2)
    REQ_QTY(0 To 10)            As Byte         '必要数         9(8)v9(2)
    ZAI_QTY(0 To 10)            As Byte         '月初在庫数     9(8)v9(2)
    MAI_QTY(0 To 10)            As Byte         '不足数         9(8)v9(2)
    ODR_QTY(0 To 10)            As Byte         '注文数         9(8)v9(2)
    SHI_QTY(0 To 10)            As Byte         '仕入残数       9(8)v9(2)
    HANSEIHIN_QTY(0 To 10)      As Byte         '半製品数       9(8)v9(2)
    
    HANSEIHIN_USE_QTY(0 To 10)  As Byte         '半製品数       9(8)v9(2)
    
    
    
    UKE_Z_QTY(0 To 10)          As Byte         '受入済み（前月以前）   9(8)v9(2)
    UKE_T_QTY(0 To 10)          As Byte         '受入済み（当月）       9(8)v9(2)
    
    
    
    
    
    
    LOT_QTY(0 To 10)            As Byte         'ロット数       9(8)v9(2)
    SECT(0 To 4)                As Byte         '仕入先
    TANKA(0 To 10)              As Byte         '発注単価       9(8)V9(2)
    NOUKI(0 To 7)               As Byte         '希望納期
    KAITO(0 To 7)               As Byte         '回答納期
    ITEM_NM(0 To 39)            As Byte         '品名
    FILLER(0 To 1)             As Byte         'Filler

End Type
'データ・バッファ
Public ODR_TP3_R            As ODR_TP3_R_Tag



'キー定義

Type KEY0_ODR_TEMP3                           'ＫＥＹ０

    USE_YM(0 To 5)              As Byte         '使用月
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    
End Type
Type KEY1_ODR_TEMP3                           'ＫＥＹ１
    
    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    
End Type

'キー・データ
Public K0_ODR_TEMP3           As KEY0_ODR_TEMP3
Public K1_ODR_TEMP3           As KEY1_ODR_TEMP3

Type ODR_TEMP3_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private ODR_TEMP3_Speck       As ODR_TEMP3_FSpeck
Private Function ODR_TEMP3_Create() As Integer
'********************************************************************
'*                                                                  *
'*              ODR_TEMP3  ＣＲＥＡＴＥ                            *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long
    
    
    ODR_TEMP3_Create = True
                                            'ODR_TEMP3 フルパス取込み
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_Str
    
    ODR_TEMP3_Speck.fs.recoleng = Len(ODR_TP3_R)      ' レコード長
    ODR_TEMP3_Speck.fs.PageSize = ODR_TEMP3_PG_SIZ          ' ページサイズ
    ODR_TEMP3_Speck.fs.idexnumb = 2                       ' インデックス数
    ODR_TEMP3_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_TEMP3_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_TEMP3_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_TEMP3_Speck.ks0.keyleng = 28                      ' キー長
    ODR_TEMP3_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP3_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP3_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    '--------------------------------------------------- キー１ ▽
    ODR_TEMP3_Speck.ks1.keypos = 7                        ' キーポジション
    ODR_TEMP3_Speck.ks1.keyleng = 22                      ' キー長
    ODR_TEMP3_Speck.ks1.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_TEMP3_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    ODR_TEMP3_Speck.ks1.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー１ △
    

    sts = BTRV(BtOpCreate, ODR_TP3_POS, ODR_TEMP3_Speck, Len(ODR_TEMP3_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_TEMP3")
        Exit Function
    End If
    
    ODR_TEMP3_Create = False

End Function

Public Function ODR_TEMP3_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_TEMP3  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim yn          As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_TEMP3_Open = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)


    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If
    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)
    
    FullPath = W_Str



    Do
        sts = BTRV(BtOpOpen, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_TEMP3_Create()      'ODR_TEMP3 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_TEMP3")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_TEMP3")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP3_Open = False
    
End Function

Public Function ODR_TEMP3_KILL() As Integer
'********************************************************************
'*
'*              所要量Ｆ  削除＆再作成（Ｏｐｅｎ）
'*
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
Dim W_Str       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_TEMP3_KILL = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_TEMP3_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_TEMP3]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    
    X_i = InStr(1, FullPath, "*") - 1
    If X_i <= 0 Then
        X_i = Len(Trim(FullPath)) - 4
    End If

    W_Str = Left(FullPath, X_i) & "_" & W_PC & ".TMP" 'Right(FullPath, 4)

    FullPath = W_Str
    
    On Error Resume Next
    Kill FullPath
    On Error GoTo 0
    
    ODR_TEMP3_KILL = False
    
End Function

Public Function ODR_TEMP3_GET(JB As String, NG As String, HG As String, Locked As Integer) As Integer
'           引数

'   JB      事業部
'   NG      内外
'   HG      子品番

'   Locked  ＧｅｔＬｏｃｋ
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_TEMP3_GET = True
    
    Call UniCode_Conv(K0_ODR_TEMP3.KO_JGYOBU, JB)       '子　事業部
    Call UniCode_Conv(K0_ODR_TEMP3.KO_NAIGAI, NG)       '子　国内外
    Call UniCode_Conv(K0_ODR_TEMP3.KO_HIN_GAI, HG)      '子品番
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), K0_ODR_TEMP3, Len(K0_ODR_TEMP3), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                'MsgBox "指定された工程がありません。"
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_TEMP3")
                Exit Function
        End Select
    Loop
    
    ODR_TEMP3_GET = False

End Function
    

Public Sub ODR_TEMP3_CLR()
    '使用月
    Call UniCode_Conv(ODR_TP3_R.USE_YM, "")
    '子　事業部
    Call UniCode_Conv(ODR_TP3_R.KO_JGYOBU, "")
    '子　国内外
    Call UniCode_Conv(ODR_TP3_R.KO_NAIGAI, "")
    '子品番
    Call UniCode_Conv(ODR_TP3_R.KO_HIN_GAI, "")
    
    '使用数         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.USE_QTY, String(UBound(ODR_TP3_R.USE_QTY) + 1, "0"))
    '必要数         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.REQ_QTY, String(UBound(ODR_TP3_R.REQ_QTY) + 1, "0"))
    '月初在庫数     9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.ZAI_QTY, String(UBound(ODR_TP3_R.ZAI_QTY) + 1, "0"))
    '不足数         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.MAI_QTY, String(UBound(ODR_TP3_R.MAI_QTY) + 1, "0"))
    '注文数         9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.ODR_QTY, String(UBound(ODR_TP3_R.ODR_QTY) + 1, "0"))
    
    
    '受入済み(前月まで分)    9(8)v9(2)  2008.05.21
    Call UniCode_Conv(ODR_TP3_R.UKE_Z_QTY, String(UBound(ODR_TP3_R.UKE_Z_QTY) + 1, "0"))
    '受入済み(当月分)    9(8)v9(2)  2008.05.21
    Call UniCode_Conv(ODR_TP3_R.UKE_T_QTY, String(UBound(ODR_TP3_R.UKE_T_QTY) + 1, "0"))
    
    
    
    '仕入残数       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.SHI_QTY, String(UBound(ODR_TP3_R.SHI_QTY) + 1, "0"))
    '半製品数       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.HANSEIHIN_QTY, String(UBound(ODR_TP3_R.HANSEIHIN_QTY) + 1, "0"))
    
    
    '半製品数       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.HANSEIHIN_USE_QTY, String(UBound(ODR_TP3_R.HANSEIHIN_USE_QTY) + 1, "0"))
    
    
    'ロット数       9(8)v9(2)
    Call UniCode_Conv(ODR_TP3_R.LOT_QTY, String(UBound(ODR_TP3_R.LOT_QTY) + 1, "0"))
    
    '仕入先
    Call UniCode_Conv(ODR_TP3_R.SECT, "")
    '発注単価       9(8)V9(2)
    Call UniCode_Conv(ODR_TP3_R.TANKA, String(UBound(ODR_TP3_R.TANKA) + 1, "0"))
    '希望納期
    Call UniCode_Conv(ODR_TP3_R.NOUKI, "")
    '回答納期
    Call UniCode_Conv(ODR_TP3_R.KAITO, "")
    '品名
    Call UniCode_Conv(ODR_TP3_R.ITEM_NM, "")
    
    Call UniCode_Conv(ODR_TP3_R.FILLER, "")
End Sub

