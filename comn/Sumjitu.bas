Attribute VB_Name = "SUMJ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              入出荷実績集計データ　ファイル定義                          *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'ファイルＩＤ
Global Const SUMJ_ID = "SUMJ"

'ページサイズ
Global Const SUMJ_PG_SIZ% = 2048

'ポジション・ブロック
Global SUMJ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SUMJREC_Tag
    JGYOBU(0 To 0) As Byte          '事業部区分
    NAIGAI(0 To 0) As Byte          '国内外
    HIN_GAI(0 To 19) As Byte        '品番（外部）
    NYUKA_QTY(0 To 7) As Byte       '入荷総数
    CHOKU_QTY(0 To 7) As Byte       '入荷直送分
    TUK_QTY(0 To 7) As Byte         '月切り出荷数
    HSP_QTY(0 To 7) As Byte         '補充スポット出荷数 (含め特売)
    BOU_QTY(0 To 7) As Byte         '貿易出荷数
    KIN_QTY(0 To 7) As Byte         '緊急出荷数
    ZAI_PURA(0 To 7) As Byte        '在訂（＋）出庫数
    ZAI_MINA(0 To 7) As Byte        '在訂（－）出庫数
    FILLER(0 To 9) As Byte          'FILLER
End Type

'データ・バッファ
Global SUMJREC As SUMJREC_Tag

'キー定義
Type KEY0_SUMJ            'ＫＥＹ０
    JGYOBU(0 To 0) As Byte          '事業部区分
    NAIGAI(0 To 0) As Byte          '国内外
    HIN_GAI(0 To 19) As Byte        '品番（外部）
End Type

'キー・データ
Global K0_SUMJ As KEY0_SUMJ

Type SUMJ_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Global SUMJ_Speck As SUMJ_FSpeck

Private Function SUMJ_Create() As Integer
'********************************************************************
'*                                                                  *
'*              入出荷実績集計データ　ＣＲＥＡＴＥ                        *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    SUMJ_Create = False
                                            '入出荷実績集計データフルパス取込み
    sts = GetIni("FILE", SUMJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SUMJ_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    SUMJ_Speck.fs.recoleng = Len(SUMJREC)     ' レコード長
    SUMJ_Speck.fs.PageSize = SUMJ_PG_SIZ      ' ページサイズ
    SUMJ_Speck.fs.idexnumb = 1                 ' インデックス数
    SUMJ_Speck.fs.fileflag = 0                 ' ファイルフラグ
    SUMJ_Speck.fs.reserve = &H0                ' 予約済み
                                                ' キー０
    SUMJ_Speck.ks0.keypos = 1                  ' キーポジション
    SUMJ_Speck.ks0.keyleng = 1 + 1 + 20        ' キー長
    SUMJ_Speck.ks0.keyflag = BtKfExt           ' キーフラグ
    SUMJ_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SUMJ_Speck.ks0.reserve = &H0               ' 予約済み

    sts = BTRV(BtOpCreate, SUMJ_POS, SUMJ_Speck, Len(SUMJ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入出荷実績集計データ")
        SUMJ_Create = True
    End If
End Function

Function SUMJ_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              入出荷実績集計データ　ＯＰＥＮ                            *
'*                                                                  *
'*      引  数:Open Mode(Btrieve参照)                               *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    SUMJ_Open = False
                                            '入出荷実績集計データフルパス取込み
    sts = GetIni("FILE", SUMJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SUMJ_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, SUMJ_POS, SUMJREC, Len(SUMJREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SUMJ_Create()        '入出荷実績集計データ作成
                If sts <> False Then
                    SUMJ_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SUMJ_POS, SUMJREC, Len(SUMJREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入出荷実績集計データ")
                    SUMJ_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "入出荷実績集計データ")
                SUMJ_Open = True
                Exit Function
        End Select
    Loop
End Function


