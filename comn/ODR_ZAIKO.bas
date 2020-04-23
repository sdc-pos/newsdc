Attribute VB_Name = "ODR_ZAIKO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              月初在庫Ｆ（WORK) ファイル定義              　　　*
'*                                                                  *
'*          CREATE 2008.08.09                                       *
'********************************************************************
'ファイルＩＤ
Public Const ODR_ZAIKO_ID$ = "ODR_ZAIKO"

'ページサイズ
Private Const ODR_ZAIKO_PG_SIZ% = 4096

'ポジション・ブロック
Public ODR_ZK_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Public Type ODR_Z_QTY_Tag
    Z_QTY(0 To 8)               As Byte         '月初在庫
    O_QTY(0 To 8)               As Byte         '追加注文数
    Y_QTY(0 To 8)               As Byte         '予備項目
End Type


Public Type ODR_ZK_R_Tag

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番
    ALL_ZAI(0 To 23)            As ODR_Z_QTY_Tag            '基準月～２４ケ月
    FILLER(0 To 29)             As Byte         'Filler

End Type
'データ・バッファ
Public ODR_ZK_R            As ODR_ZK_R_Tag



'キー定義

Type KEY0_ODR_ZAIKO                           'ＫＥＹ０

    KO_JGYOBU(0 To 0)           As Byte         '子　事業部
    KO_NAIGAI(0 To 0)           As Byte         '子　国内外
    KO_HIN_GAI(0 To 19)         As Byte         '子品番

End Type

'キー・データ
Public K0_ODR_ZK            As KEY0_ODR_ZAIKO


Type ODR_ZAIKO_FSpeck
    fs                      As BtFileSpeck  ' ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0                     As BtKeySpeck   ' ｷｰ ｽﾍﾟｯｸ構造体

End Type

Private ODR_ZAIKO_Speck       As ODR_ZAIKO_FSpeck
Private Function ODR_ZAIKO_Create() As Integer
'*******************************************************************
'*                                                                 *
'*              ODR_ZAIKO  ＣＲＥＡＴＥ                             *
'*                                                                 *
'*      引  数:なし                                                 *
'*      戻り値:false 正常                                           *
'*             true  異常                                          *
'*                                                                 *
'*******************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    ODR_ZAIKO_Create = True
                                            'ODR_ZAIKO フルパス取込み
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    ODR_ZAIKO_Speck.fs.recoleng = Len(ODR_ZK_R)      ' レコード長
    ODR_ZAIKO_Speck.fs.PageSize = ODR_ZAIKO_PG_SIZ          ' ページサイズ
    ODR_ZAIKO_Speck.fs.idexnumb = 1                       ' インデックス数
    ODR_ZAIKO_Speck.fs.fileflag = 0                       ' ファイルフラグ
    ODR_ZAIKO_Speck.fs.reserve = &H0                      ' 予約済み
    '--------------------------------------------------- キー０ ▽
    ODR_ZAIKO_Speck.ks0.keypos = 1                        ' キーポジション
    ODR_ZAIKO_Speck.ks0.keyleng = 22                      ' キー長
    ODR_ZAIKO_Speck.ks0.keyflag = BtKfChg + BtKfDup + BtKfExt      ' キーフラグ
    ODR_ZAIKO_Speck.ks0.keytype = Chr(BtKtString)         ' キータイプ
    ODR_ZAIKO_Speck.ks0.reserve = &H0                     ' 予約済み
    '--------------------------------------------------- キー０ △
    
    

    sts = BTRV(BtOpCreate, ODR_ZK_POS, ODR_ZAIKO_Speck, Len(ODR_ZAIKO_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ODR_ZAIKO")
        Exit Function
    End If
    
    ODR_ZAIKO_Create = False

End Function

Public Function ODR_ZAIKO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ODR_ZAIKO  ＯＰＥＮ
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
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long

    ODR_ZAIKO_Open = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                yn = MsgBox("他で使用中です！<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function

            Case BtErrFileNotFound
                sts = ODR_ZAIKO_Create()            'ODR_ZAIKO 作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ODR_ZAIKO")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ODR_ZAIKO")
                Exit Function
        End Select
    Loop
    
    ODR_ZAIKO_Open = False
    
End Function

Public Function ODR_ZAIKO_KILL() As Integer
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
Dim W_STR       As String
Dim W_PC        As String
Dim X_i         As Long
Dim X_j         As Long

    ODR_ZAIKO_KILL = True
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_ZAIKO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_ZAIKO]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)


    
    Kill FullPath
    
    ODR_ZAIKO_KILL = False
    
End Function

Public Function ODR_ZAIKO_GET(JB As String, NG As String, HG As String, _
                      Locked As Integer) As Integer
'           引数

'   JB      事業部
'   NG      内外
'   HG      子品番

'   Locked  ＧｅｔＬｏｃｋ
    
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

    ODR_ZAIKO_GET = True
    
    Call UniCode_Conv(K0_ODR_ZK.KO_JGYOBU, JB)       '子　事業部
    Call UniCode_Conv(K0_ODR_ZK.KO_NAIGAI, NG)       '子　国内外
    Call UniCode_Conv(K0_ODR_ZK.KO_HIN_GAI, HG)      '子品番
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<ODR_ZAIKO>" & Chr(13) & Chr(10) & _
                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "ODR_ZAIKO")
                Exit Function
        End Select
    Loop
    
    ODR_ZAIKO_GET = False

End Function
    

Public Sub ODR_ZAIKO_CLR()
Dim X_i As Integer

    '子　事業部
    Call UniCode_Conv(ODR_ZK_R.KO_JGYOBU, "")
    '子　国内外
    Call UniCode_Conv(ODR_ZK_R.KO_NAIGAI, "")
    '子品番
    Call UniCode_Conv(ODR_ZK_R.KO_HIN_GAI, "")
    
    For X_i = 0 To UBound(ODR_ZK_R.ALL_ZAI)
                                            '在庫数     9(5)v9(2)
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).Z_QTY) + 1, "0"))
                                            '注文数     9(5)v9(2)
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).O_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).O_QTY) + 1, "0"))
                                            '予備
        Call UniCode_Conv(ODR_ZK_R.ALL_ZAI(X_i).Y_QTY, String(UBound(ODR_ZK_R.ALL_ZAI(X_i).Y_QTY) + 1, "0"))
    Next X_i
    
    Call UniCode_Conv(ODR_ZK_R.FILLER, "")
End Sub

