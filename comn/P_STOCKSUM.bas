Attribute VB_Name = "P_STOCKSUM"
Option Explicit

'********************************************************************
'*
'*              資材棚卸集計ﾃﾞｰﾀ  ファイル定義
'*
'*          CREATE 2006.02.15
'********************************************************************
'ファイルＩＤ
Public Const P_STOCKSUM_ID$ = "P_STOCKSUM"

'ページサイズ
Private Const P_STOCKSUM_PG_SIZ% = 1024

'ポジション・ブロック
Public P_STOCKSUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義


Public Type P_STOCKSUM_REC_Tag
    G_SYUSHI(0 To 2)            As Byte         '収支単位
    ZEN_ZAIKO_KIN(0 To 10)      As Byte         '前月在庫金額

    NYUKO_KIN(0 To 10)          As Byte         '当月入庫金額
    SYUKO_KIN(0 To 10)          As Byte         '当月出庫金額
    ZAIKO_KIN(0 To 10)          As Byte         '現在庫金額
    FILLER(0 To 16)             As Byte         '


End Type
'データ・バッファ
Public P_STOCKSUM_REC          As P_STOCKSUM_REC_Tag

'キー定義
    
Public Type KEY0_P_STOCKSUM                    'ＫＥＹ０
    G_SYUSHI(0 To 2)            As Byte         '収支単位
End Type
    
    
'キー・データ
Public K0_P_STOCKSUM        As KEY0_P_STOCKSUM

Type P_STOCKSUM_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_STOCKSUM_Speck       As P_STOCKSUM_FSpeck
Private Function P_STOCKSUM_Create() As Integer
'********************************************************************
'*
'*              資材棚卸集計ﾃﾞｰﾀ  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*      収支毎にファイル名を分ける  2007.11.13
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Long     '2007.11.13




    P_STOCKSUM_Create = True
                                            '資材棚卸集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]読み込みエラー")
        Exit Function
    End If



    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13
   
    
    P_STOCKSUM_Speck.fs.recoleng = Len(P_STOCKSUM_REC)        ' レコード長
    P_STOCKSUM_Speck.fs.PageSize = P_STOCKSUM_PG_SIZ          ' ページサイズ
    P_STOCKSUM_Speck.fs.idexnumb = 1                       ' インデックス数
    P_STOCKSUM_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_STOCKSUM_Speck.fs.reserve = &H0                      ' 予約済み
    
    '--------------------------------------------------- キー０ ▽
    P_STOCKSUM_Speck.ks0.keypos = 1                        ' キーポジション
    P_STOCKSUM_Speck.ks0.keyleng = 3                       ' キー長
    P_STOCKSUM_Speck.ks0.keyflag = BtKfExt                 ' キーフラグ
    P_STOCKSUM_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_STOCKSUM_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    
    
    sts = BTRV(BtOpCreate, P_STOCKSUM_POS, P_STOCKSUM_Speck, Len(P_STOCKSUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "資材棚卸集計ﾃﾞｰﾀ")
        Exit Function
    End If
    
    P_STOCKSUM_Create = False

End Function

Public Function P_STOCKSUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              資材棚卸集計ﾃﾞｰﾀ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*      収支毎にファイル名を分ける  2007.11.13
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    P_STOCKSUM_Open = True
                                            '資材棚卸集計ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]読み込みエラー")
        Exit Function
    End If
    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13

    Do
        sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_STOCKSUM_Create()   '資材棚卸集計ﾃﾞｰﾀ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "資材棚卸集計ﾃﾞｰﾀ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "資材棚卸集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    P_STOCKSUM_Open = False

End Function

