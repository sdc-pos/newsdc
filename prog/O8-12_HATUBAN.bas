Attribute VB_Name = "O_HATUBN"
Option Explicit
'********************************************************************
'*
'*              発番マスタ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const O_HATUBAN_ID$ = "O_HATUBAN"

'ページサイズ
Public Const O_HATUBAN_PG_SIZ% = 512

'ポジション・ブロック
Public O_HATUBAN_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type O_HATUBANREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NYK_KBN(0 To 0)         As Byte         '入荷伝票№区分
    NYK_DEN_NO(0 To 4)      As Byte         '次入荷伝票№
    SYK_KBN(0 To 0)         As Byte         '出荷伝票№区分
    SYK_DEN_NO(0 To 4)      As Byte         '次出荷伝票№
    NYK_ID_KBN(0 To 0)      As Byte         '入荷ID№区分
    NYK_ID_NO(0 To 7)       As Byte         '次入荷ID№
    SYK_ID_KBN(0 To 0)      As Byte         '出荷ID№区分
    SYK_ID_NO(0 To 6)       As Byte         '次出荷ID№
    FILLER(0 To 11)         As Byte         'FILLER
End Type

'データ・バッファ
Public O_HATUBANREC           As O_HATUBANREC_Tag

'キー定義
Type KEY0_O_HATUBAN            'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
End Type

'キー・データ
Public K0_O_HATUBAN           As KEY0_O_HATUBAN

Type O_HATUBAN_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private O_HATUBAN_Speck As O_HATUBAN_FSpeck

Private Function O_HATUBAN_Create() As Integer
'********************************************************************
'*
'*              発番マスタ　ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    O_HATUBAN_Create = True
                                            '発番マスタフルパス取込み
    sts = GetIni("FILE", O_HATUBAN_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_HATUBAN]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    O_HATUBAN_Speck.fs.recoleng = Len(O_HATUBANREC)     ' レコード長
    O_HATUBAN_Speck.fs.PageSize = O_HATUBAN_PG_SIZ      ' ページサイズ
    O_HATUBAN_Speck.fs.idexnumb = 1                   ' インデックス数
    O_HATUBAN_Speck.fs.fileflag = 0                   ' ファイルフラグ
    O_HATUBAN_Speck.fs.reserve = &H0                  ' 予約済み
                                                    ' キー０
    O_HATUBAN_Speck.ks0.keypos = 1                    ' キーポジション
    O_HATUBAN_Speck.ks0.keyleng = 1                   ' キー長
    O_HATUBAN_Speck.ks0.keyflag = BtKfExt             ' キーフラグ
    O_HATUBAN_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    O_HATUBAN_Speck.ks0.reserve = &H0                 ' 予約済み

    sts = BTRV(BtOpCreate, O_HATUBAN_POS, O_HATUBAN_Speck, Len(O_HATUBAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "発番マスタ")
        Exit Function
    End If

    O_HATUBAN_Create = False

End Function

Public Function O_HATUBAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              発番マスタ　ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    O_HATUBAN_Open = True
                                            '発番マスタフルパス取込み
    sts = GetIni("FILE", O_HATUBAN_ID, "CONV200605", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [O_HATUBAN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = O_HATUBAN_Create()        '発番マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "発番マスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "発番マスタ")
                Exit Function
        End Select
    Loop

    O_HATUBAN_Open = False

End Function
