Attribute VB_Name = "P_CLASS"
Option Explicit
'********************************************************************
'*                                                                  *
'*              クラスマスタ  ファイル定義                          *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'ファイルＩＤ
Public Const P_CLASS_ID$ = "P_CLASS"

'ページサイズ
Private Const P_CLASS_PG_SIZ% = 512

'ポジション・ブロック
Public P_CLASS_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type P_CLASSREC_Tag
    
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    CLASS_CODE(0 To 19)     As Byte         'クラス（品番）
    CLASS_NAME(0 To 49)     As Byte         '呼び名
    TANKA(0 To 10)          As Byte         '商品化価格 9(8)V99
    KOUSU(0 To 6)           As Byte         '工数 999V999
    KOURYOU(0 To 10)        As Byte         '工料 9(8)V99
    ETC(0 To 10)            As Byte         'その他
'''2007.01.11    FILLER(0 To 252)        As Byte         'Filler
    URI_KOURYOU(0 To 10)    As Byte         '工料 9(8)V99   2007.01.11
    FILLER(0 To 241)        As Byte         'Filler         2007.01.11
    
    
    
    
    
    
    UPD_TANTO(0 To 4)       As Byte         '更新　担当者
    UPD_DATETIME(0 To 13)   As Byte         '更新　日時

End Type
'データ・バッファ
Public P_CLASSREC           As P_CLASSREC_Tag

'キー定義

Type KEY0_P_CLASS                           'ＫＥＹ０
    SHIMUKE_CODE(0 To 1)    As Byte         '仕向け先
    CLASS_CODE(0 To 19)     As Byte         'クラス（品番）
End Type
    
'キー・データ
Public K0_P_CLASS           As KEY0_P_CLASS

Type P_CLASS_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
    ks1                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private P_CLASS_Speck       As P_CLASS_FSpeck
Private Function P_CLASS_Create() As Integer
'********************************************************************
'*                                                                  *
'*              クラスマスタ  ＣＲＥＡＴＥ                          *
'*                                                                  *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_CLASS_Create = True
                                            'クラスマスタフルパス取込み
    sts = GetIni("FILE", P_CLASS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CLASS]読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_CLASS_Speck.fs.recoleng = Len(P_CLASSREC)         ' レコード長
    P_CLASS_Speck.fs.PageSize = P_CLASS_PG_SIZ          ' ページサイズ
    P_CLASS_Speck.fs.idexnumb = 1                       ' インデックス数
    P_CLASS_Speck.fs.fileflag = 0                       ' ファイルフラグ
    P_CLASS_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    P_CLASS_Speck.ks0.keypos = 1                        ' キーポジション
    P_CLASS_Speck.ks0.keyleng = 2                       ' キー長
    P_CLASS_Speck.ks0.keyflag = BtKfExt + BtKfSeg       ' キーフラグ
    P_CLASS_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    P_CLASS_Speck.ks0.reserve = &H0                     ' 予約済み
    
    P_CLASS_Speck.ks1.keypos = 3                        ' キーポジション
    P_CLASS_Speck.ks1.keyleng = 20                      ' キー長
    P_CLASS_Speck.ks1.keyflag = BtKfExt                 ' キーフラグ
    P_CLASS_Speck.ks1.keytype = Chr(BtKtString)         ' キータイプ
    P_CLASS_Speck.ks1.reserve = &H0                     ' 予約済み
    
    '--------------------------------------------------- キー０ △
    sts = BTRV(BtOpCreate, P_CLASS_POS, P_CLASS_Speck, Len(P_CLASS_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "クラスマスタ")
        Exit Function
    End If
    
    P_CLASS_Create = False

End Function

Public Function P_Class_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              クラスマスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_Class_Open = True
                                            'クラスマスタフルパス取込み
    sts = GetIni("FILE", P_CLASS_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_CLASS]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_CLASS_Create()      'クラスマスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "クラスマスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "クラスマスタ")
                Exit Function
        End Select
    Loop
    
    P_Class_Open = False

End Function
