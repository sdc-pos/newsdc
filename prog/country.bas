Attribute VB_Name = "country"
Option Explicit
'********************************************************************
'*
'*              Countryマスタ ファイル定義
'*
'*          CREATE 2010.09.01
'********************************************************************
'ファイルＩＤ
Public Const Country_ID = "Country"

'ページサイズ
Public Const Country_PG_SIZ% = 4096

'ポジション・ブロック
Public Country_POS  As POSBLK
'=
'=
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type CountryREC_Tag
    CountryCode(0 To 2)     As Byte     '国コード
    CountryName(0 To 19)    As Byte     '国名１
    CountryName2(0 To 19)   As Byte     '国名２
    

End Type
'データ・バッファ
Public CountryREC           As CountryREC_Tag


'キー定義
Type KEY0_Country                       'ＫＥＹ０
    CountryCode(0 To 2)     As Byte     '国コード
End Type



'キー・データ
Public K0_Country           As KEY0_Country

Private Type Country_FSpeck
    fs  As BtFileSpeck              ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0 As BtKeySpeck               ' ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private Country_Speck       As Country_FSpeck
Private Function Country_Create() As Integer
'********************************************************************
'*
'*              Countryファイル  ＣＲＥＡＴＥ
'*
'*      引  数:なし
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2010.09.01
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    Country_Create = True
                                            'Countryファイルフルパス取込み
    sts = GetIni("FILE", Country_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[Country] 読み込みエラー")
        Exit Function
    End If

    FullPath = RTrim$(c)

    Country_Speck.fs.recoleng = Len(CountryREC)     ' レコード長
    Country_Speck.fs.PageSize = Country_PG_SIZ      ' ページサイズ
    Country_Speck.fs.idexnumb = 1                   ' インデックス数
    Country_Speck.fs.fileflag = 0                   ' ファイルフラグ
    Country_Speck.fs.reserve = &H0                  ' 予約済み

'---------------------------------------------------' キー０
    Country_Speck.ks0.keypos = 1                    ' キーポジション
    Country_Speck.ks0.keyleng = 3                   ' キー長
                                                    ' キーフラグ
    Country_Speck.ks0.keyflag = BtKfExt + BtKfChg
    Country_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    Country_Speck.ks0.reserve = &H0                 ' 予約済み

    
    
    sts = BTRV(BtOpCreate, Country_POS, Country_Speck, Len(Country_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Countryマスタ")
        Exit Function
    End If

    Country_Create = False

End Function

Function Country_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              Countryマスタ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2010.09.01
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    Country_Open = True
                                            'Countryファイルフルパス取込み
    sts = GetIni("FILE", Country_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, Country_POS, CountryREC, Len(CountryREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = Country_Create()        'Countryファイル作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, Country_POS, CountryREC, Len(CountryREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Countryマスタ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Countryマスタ")
                Exit Function
        End Select
    Loop
    Country_Open = False
End Function

