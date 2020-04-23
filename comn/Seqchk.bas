Attribute VB_Name = "SEQCK"
Option Explicit
'********************************************************************
'*                                                                  *
'*              予定取込みチェック ファイル定義                       *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'ファイルＩＤ
Global Const SEQCK_ID = "SEQCK"

'ページサイズ
Global Const SEQCK_PG_SIZ% = 512

'ポジション・ブロック
Global SEQCK_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type SEQCKREC_Tag
    JGYOBU(0 To 0)      As Byte     '事業部区分
    SEQ_MODE(0 To 0)    As Byte     '取り込み区分
    LAST_TXTNO(0 To 8)  As Byte     '最終テキスト№
    LAST_GET_DT(0 To 7) As Byte     '最終取込み日付
    LAST_GET_TM(0 To 5) As Byte     '最終取込み時刻
End Type

'データ・バッファ
Global SEQCKREC         As SEQCKREC_Tag
'キー定義

Type KEY0_SEQCK            'ＫＥＹ０
    JGYOBU(0 To 0)      As Byte     '事業部区分
    SEQ_MODE(0 To 0)    As Byte     '取り込み区分
End Type
    
'キー・データ
Global K0_SEQCK         As KEY0_SEQCK

Type SEQCK_FSpeck
    fs      As BtFileSpeck          ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck           ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Global SEQCK_Speck As SEQCK_FSpeck
Private Function SEQCK_Create() As Integer
'********************************************************************
'*                                                                  *
'*              予定取込みチェック  ＣＲＥＡＴＥ                    *
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

    SEQCK_Create = False
                                            '予定取込みチェックフルパス取込み
    sts = GetIni("FILE", SEQCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SEQCK_Create = True
        Exit Function
    End If

    FullPath = RTrim$(c)

    SEQCK_Speck.fs.recoleng = Len(SEQCKREC)     ' レコード長
    SEQCK_Speck.fs.PageSize = SEQCK_PG_SIZ      ' ページサイズ
    SEQCK_Speck.fs.idexnumb = 1                 ' インデックス数
    SEQCK_Speck.fs.fileflag = 0                 ' ファイルフラグ
    SEQCK_Speck.fs.reserve = &H0                ' 予約済み
                                                ' キー０
    SEQCK_Speck.ks0.keypos = 1                  ' キーポジション
    SEQCK_Speck.ks0.keyleng = 1 + 1             ' キー長
    SEQCK_Speck.ks0.keyflag = BtKfExt           ' キーフラグ
    SEQCK_Speck.ks0.keytype = Chr(BtKtString)   ' キータイプ
    SEQCK_Speck.ks0.reserve = &H0               ' 予約済み

    sts = BTRV(BtOpCreate, SEQCK_POS, SEQCK_Speck, Len(SEQCK_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "予定取込みチェック")
        SEQCK_Create = True
    End If
End Function
Function SEQCK_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              予定取込みチェック  ＯＰＥＮ                        *
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

    SEQCK_Open = False
                                            '予定取込みチェックフルパス取込み
    sts = GetIni("FILE", SEQCK_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        SEQCK_Open = True
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, SEQCK_POS, SEQCKREC, Len(SEQCKREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Function
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SEQCK_Create()        '予定取込みチェック作成
                If sts <> False Then
                    SEQCK_Open = True
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SEQCK_POS, SEQCKREC, Len(SEQCKREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "予定取込みチェック")
                    SEQCK_Open = True
                End If
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "予定取込みチェック")
                SEQCK_Open = True
                Exit Function
        End Select
    Loop
End Function

