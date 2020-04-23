Attribute VB_Name = "PRGF"
Option Explicit
'********************************************************************
'*                                                                  *
'*              プログラムチェックファイル定義                      *
'*                                                                  *
'*          CREATE 1997.06.02  S.Shibano                            *
'********************************************************************
'ファイルＩＤ
Global Const PRGF_ID = "PRGF"

'ページサイズ
Global Const PRGF_PG_SIZ% = 512
'ファイルパス
Global PRGFPath As String
'レコード定義
Type PRGF
    PROG_ID(0 To 7) As Byte         'プログラムＩＤ
    END_CTL(0 To 0) As Byte         '終了制御有無
    START_DT(0 To 7) As Byte        '開始日付
    START_TM(0 To 5) As Byte        '開始時刻
    FILLER(0 To 6) As Byte          'FILLER
End Type
'ＫＥＹ０
Type KEY0_PRGF_Tag
    PROG_ID(0 To 7) As Byte         'プログラムＩＤ
End Type
'ＫＥＹ１
Type KEY1_PRGF_Tag
    END_CTL(0 To 0) As Byte         '終了制御有無
    PROG_ID(0 To 7) As Byte         'プログラムＩＤ
End Type

Global PRGFRec As PRGF

Global K0_PRGF As KEY0_PRGF_Tag
Global K1_PRGF As KEY1_PRGF_Tag

Global PRGF_Pos As POSBLK
    
Type PRGF_FSpeck
    fs As BtFileSpeck                 ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck                 ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1 As BtKeySpeck
    ks2 As BtKeySpeck
End Type

Global PRGF_Speck As PRGF_FSpeck



'****************************************************
'*      同一プログラム実行チェックファイル作成      *
'*  引  数: なし                                    *
'*                                                  *
'*  戻り値: false   正常終了                        *
'*          true    異常終了                        *
'*          CREATE 1997.06.06  S.Shibano            *
'****************************************************
Private Function PRGF_Create() As Integer
Dim sts As Integer
Dim c As String * 128
Dim messge As String
    
    PRGF_Create = False
    
    sts = GetIni("FILE", PRGF_ID, "SYS", c)
    If sts <> False Then
        messge = "SYS.INI 読込みエラー"
        Call Log_Out(LOG_F, messge)
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        PRGF_Create = True
        Exit Function
    End If
    PRGFPath = RTrim(c)
    PRGF_Speck.fs.recoleng = Len(PRGFRec)           ' レコード長
    PRGF_Speck.fs.PageSize = PRGF_PG_SIZ            ' ページサイズ
    PRGF_Speck.fs.idexnumb = 2                      ' インデックス数
    PRGF_Speck.fs.fileflag = 0                      ' ファイルフラグ
    PRGF_Speck.fs.reserve = &H0                     ' 予約済み
                                                    ' キー０
    PRGF_Speck.ks0.keypos = 1                       ' キーポジション
    PRGF_Speck.ks0.keyleng = 8                      ' キー長
    PRGF_Speck.ks0.keyflag = BtKfExt                ' キーフラグ
    PRGF_Speck.ks0.keytype = Chr(BtKtString)        ' キータイプ
    PRGF_Speck.ks0.reserve = &H0                    ' 予約済み
                                                    ' キー１
    PRGF_Speck.ks1.keypos = 9                       ' キーポジション
    PRGF_Speck.ks1.keyleng = 1                      ' キー長
    PRGF_Speck.ks1.keyflag = BtKfExt + BtKfSeg      ' キーフラグ
    PRGF_Speck.ks1.keytype = Chr(BtKtString)        ' キータイプ
    PRGF_Speck.ks1.reserve = &H0                    ' 予約済み
                                                    ' キー１
    PRGF_Speck.ks2.keypos = 1                       ' キーポジション
    PRGF_Speck.ks2.keyleng = 8                      ' キー長
    PRGF_Speck.ks2.keyflag = BtKfExt                ' キーフラグ
    PRGF_Speck.ks2.keytype = Chr(BtKtString)        ' キータイプ
    PRGF_Speck.ks2.reserve = &H0                    ' 予約済み

    sts = BTRV(BtOpCreate, PRGF_Pos, PRGF_Speck, Len(PRGF_Speck), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "起動プロセスチェック")
        PRGF_Create = True
        Exit Function
    End If
    sts = BTRV(BtOpOpen, PRGF_Pos, PRGFRec, Len(PRGFRec), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        Call File_Error(sts, BtOpOpen, "起動プロセスチェック")
        PRGF_Create = True
        Exit Function
    End If
End Function

'****************************************************
'*      同一プログラム実行チェックファイルＯＰＥＮ  *
'*  引  数: Open Mode(Btrieve参照)                  *
'*                                                  *
'*  戻り値: false   正常終了                        *
'*          true    異常終了                        *
'*          CREATE 1997.06.06  S.Shibano            *
'****************************************************
Function PRGF_Open(Mode As Integer) As Integer

Dim c As String * 128
Dim messge As String
Dim sts As Integer

    PRGF_Open = False

    sts = GetIni("FILE", PRGF_ID, "SYS", c)
    If sts <> False Then
        messge = "SYS.INI 読込みエラー"
        Call Log_Out(LOG_F, messge)
        PRGF_Open = True
        Exit Function
    End If
    PRGFPath = RTrim(c)
    sts = BTRV(BtOpOpen, PRGF_Pos, PRGFRec, Len(PRGFRec), ByVal PRGFPath, Len(PRGFPath), 0)
    If sts Then
        If sts = BtErrFileNotFound Then
            sts = PRGF_Create()
            If sts Then
                PRGF_Open = True
                Exit Function
            End If
        Else
            Call File_Error(sts, BtOpOpen, "起動プロセスチェック")
            PRGF_Open = True
            Exit Function
        End If
    End If
End Function
