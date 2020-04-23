Attribute VB_Name = "MEIJ"
Option Explicit
'********************************************************************
'*                                                                  *
'*              入出荷明細データ　ファイル定義                        *
'*                                                                  *
'*          CREATE 2001.05.15                                       *
'********************************************************************
'ファイルＩＤ
Global Const MEIJ_ID = "MEIJ"

'ページサイズ
Global Const MEIJ_PG_SIZ% = 2048

'ポジション・ブロック
Global MEIJ_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                              *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type MEIJREC_Tag
    IO_KBN(0 To 0)      As Byte     '入荷／出荷区分
    DEN_DT(0 To 7)      As Byte     '伝票日付
    CYU_KBN(0 To 0)     As Byte     '注文区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
    JITU_QTY(0 To 8)    As Byte     '実績数
    FILLER(0 To 7)     As Byte     'FILLER
End Type

'データ・バッファ
Public MEIJREC As MEIJREC_Tag

'キー定義
Type KEY0_MEIJ            'ＫＥＹ０
    IO_KBN(0 To 0)      As Byte     '入荷／出荷区分
    DEN_DT(0 To 7)      As Byte     '伝票日付
    CYU_KBN(0 To 0)     As Byte     '注文区分
    NAIGAI(0 To 0)      As Byte     '国内外
    HIN_GAI(0 To 19)    As Byte     '品番（外部）
End Type

'キー・データ
Public K0_MEIJ As KEY0_MEIJ

Type MEIJ_FSpeck
    fs As BtFileSpeck               ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Global MEIJ_Speck As MEIJ_FSpeck

Private Function MEIJ_Create() As Integer
'********************************************************************
'*
'*              入出荷明細データ　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.05.15
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    MEIJ_Create = False
                                            '入出荷実績集計データフルパス取込み
    sts = GetIni("FILE", MEIJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        MEIJ_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    MEIJ_Speck.fs.recoleng = Len(MEIJREC)       ' レコード長
    MEIJ_Speck.fs.PageSize = MEIJ_PG_SIZ        ' ページサイズ
    MEIJ_Speck.fs.idexnumb = 1                  ' インデックス数
    MEIJ_Speck.fs.fileflag = 0                  ' ファイルフラグ
    MEIJ_Speck.fs.reserve = &H0                 ' 予約済み
                                                ' キー０
    MEIJ_Speck.ks0.keypos = 1                   ' キーポジション
    MEIJ_Speck.ks0.keyleng = 1 + 8 + 1 + 1 + 20 ' キー長
    MEIJ_Speck.ks0.keyflag = BtKfExt            ' キーフラグ
    MEIJ_Speck.ks0.keytype = Chr(BtKtString)    ' キータイプ
    MEIJ_Speck.ks0.reserve = &H0                ' 予約済み

    sts = BTRV(BtOpCreate, MEIJ_POS, MEIJ_Speck, Len(MEIJ_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "入出荷明細データ")
        MEIJ_Create = True
    End If
End Function

Function MEIJ_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              入出荷明細データ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.05.15
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    MEIJ_Open = False
                                            '入出荷明細データフルパス取込み
    sts = GetIni("FILE", MEIJ_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        MEIJ_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, MEIJ_POS, MEIJREC, Len(MEIJREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = MEIJ_Create()        '入出荷実績集計データ作成
                If sts <> False Then
                    MEIJ_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, MEIJ_POS, MEIJREC, Len(MEIJREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "入出荷明細データ")
                    MEIJ_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "入出荷明細データ")
                MEIJ_Open = True
                Exit Function
        End Select
    Loop
End Function


