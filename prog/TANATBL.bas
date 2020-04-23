Attribute VB_Name = "TANATBL"
Option Explicit
'********************************************************************
'*
'*              棚番読替えテーブル  ファイル定義
'*
'*          CREATE 2001.06.13
'********************************************************************
'ファイルＩＤ
Global Const TANATBL_ID = "TANATBL"

'ページサイズ
Global Const TANATBL_PG_SIZ% = 512

'ポジション・ブロック
Global TANATBL_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type TANATBLREC_Tag
    HOST_TANA(0 To 7)   As Byte
    FILLER(0 To 0)      As Byte
    POS_TANA(0 To 7)    As Byte
End Type

'データ・バッファ
Global TANATBLREC As TANATBLREC_Tag

'キー定義

Type KEY0_TANATBL                 'ＫＥＹ０
    HOST_TANA(0 To 7)   As Byte
End Type

'キー・データ
Global K0_TANATBL As KEY0_TANATBL

Type TANATBL_FSpeck
    fs As BtFileSpeck           ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Global TANATBL_Speck As TANATBL_FSpeck

Private Function TANATBL_Create() As Integer
'********************************************************************
'*
'*              棚番読替えテーブル  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.02.14
'********************************************************************
Dim sts As Integer
Dim FullPath As String
Dim c As String * 128

    TANATBL_Create = False
                                            '棚番読替えテーブルフルパス取込み
    sts = GetIni("FILE", TANATBL_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        TANATBL_Create = True
        Exit Function
    End If
    
    FullPath = RTrim$(c)
    
    TANATBL_Speck.fs.recoleng = Len(TANATBLREC)         ' レコード長
    TANATBL_Speck.fs.PageSize = TANATBL_PG_SIZ          ' ページサイズ
    TANATBL_Speck.fs.idexnumb = 1                   ' インデックス数
    TANATBL_Speck.fs.fileflag = 0                   ' ファイルフラグ
    TANATBL_Speck.fs.reserve = &H0                  ' 予約済み
                                                ' キー０
    TANATBL_Speck.ks0.keypos = 1                    ' キーポジション
    TANATBL_Speck.ks0.keyleng = 8                   ' キー長
    TANATBL_Speck.ks0.keyflag = BtKfExt             ' キーフラグ
    TANATBL_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    TANATBL_Speck.ks0.reserve = &H0                 ' 予約済み
    sts = BTRV(BtOpCreate, TANATBL_POS, TANATBL_Speck, Len(TANATBL_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "棚番読替えテーブル")
        TANATBL_Create = True
    End If
End Function

Function TANATBL_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              棚番読替えテーブル  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2001.06.13
'********************************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String
    
    TANATBL_Open = False
                                            '棚番読替えテーブルフルパス取込み
    sts = GetIni("FILE", TANATBL_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        TANATBL_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, TANATBL_POS, TANATBLREC, Len(TANATBLREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = TANATBL_Create()        '棚番読替えテーブル作成
                If sts <> False Then
                    TANATBL_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, TANATBL_POS, TANATBLREC, Len(TANATBLREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "棚番読替えテーブル")
                    TANATBL_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "棚番読替えテーブル")
                TANATBL_Open = True
                Exit Function
        End Select
    Loop
End Function
