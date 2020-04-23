Attribute VB_Name = "HATUBN"
Option Explicit
'********************************************************************
'*
'*              発番マスタ　ファイル定義
'*
'********************************************************************
'ファイルＩＤ
Public Const HATUBAN_ID$ = "HATUBAN"

'ページサイズ
Public Const HATUBAN_PG_SIZ% = 512

'ポジション・ブロック
Public HATUBAN_POS As POSBLK
'********************************************************************
'*
'*                           構造体定義
'*
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義
Type HATUBANREC_Tag
    JGYOBU(0 To 0)          As Byte         '事業部区分
    NYK_KBN(0 To 0)         As Byte         '入荷伝票№区分
    NYK_DEN_NO(0 To 4)      As Byte         '次入荷伝票№
    SYK_KBN(0 To 0)         As Byte         '出荷伝票№区分
    SYK_DEN_NO(0 To 4)      As Byte         '次出荷伝票№
    NYK_ID_KBN(0 To 0)      As Byte         '入荷ID№区分
    NYK_ID_NO(0 To 7)       As Byte         '次入荷ID№
    SYK_ID_KBN(0 To 0)      As Byte         '出荷ID№区分
    SYK_ID_NO(0 To 10)      As Byte         '次出荷ID№         2006.05.23 7-->11

    OPC_ID_KBN(0 To 0)      As Byte         '大阪PCID№区分     2006.12.11
    OPC_ID_NO(0 To 5)       As Byte         '大阪PC次出荷ID№   2006.12.11

    OPC_DEN_KBN(0 To 0)     As Byte         '大阪PC伝票№区分   2006.12.11
    OPC_DEN_NO(0 To 5)      As Byte         '大阪PC伝票№       2006.12.11

    OPC_SYU_NO(0 To 11)     As Byte         '大阪PC出庫表№     2007.03.15


    FILLER(0 To 19)         As Byte         'FILLER             2006.12.11
End Type

'データ・バッファ
Public HATUBANREC           As HATUBANREC_Tag

'キー定義
Type KEY0_HATUBAN            'ＫＥＹ０
    JGYOBU(0 To 0)          As Byte         '事業部区分
End Type

'キー・データ
Public K0_HATUBAN           As KEY0_HATUBAN

Type HATUBAN_FSpeck
    fs      As BtFileSpeck                  'ﾌｧｲﾙ ｽﾍﾟｯｸ構造体
    ks0     As BtKeySpeck                   'ｷｰ ｽﾍﾟｯｸ構造体
End Type

Private HATUBAN_Speck As HATUBAN_FSpeck

Private Function HATUBAN_Create() As Integer
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

    HATUBAN_Create = True
                                            '発番マスタフルパス取込み
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HATUBAN]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    HATUBAN_Speck.fs.recoleng = Len(HATUBANREC)     ' レコード長
    HATUBAN_Speck.fs.PageSize = HATUBAN_PG_SIZ      ' ページサイズ
    HATUBAN_Speck.fs.idexnumb = 1                   ' インデックス数
    HATUBAN_Speck.fs.fileflag = 0                   ' ファイルフラグ
    HATUBAN_Speck.fs.reserve = &H0                  ' 予約済み
                                                    ' キー０
    HATUBAN_Speck.ks0.keypos = 1                    ' キーポジション
    HATUBAN_Speck.ks0.keyleng = 1                   ' キー長
    HATUBAN_Speck.ks0.keyflag = BtKfExt             ' キーフラグ
    HATUBAN_Speck.ks0.keytype = Chr(BtKtString)     ' キータイプ
    HATUBAN_Speck.ks0.reserve = &H0                 ' 予約済み

    sts = BTRV(BtOpCreate, HATUBAN_POS, HATUBAN_Speck, Len(HATUBAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "発番マスタ")
        Exit Function
    End If

    HATUBAN_Create = False

End Function

Public Function HATUBAN_Open(Mode As Integer) As Integer
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
    
    HATUBAN_Open = True
                                            '発番マスタフルパス取込み
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HATUBAN]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HATUBAN_Create()        '発番マスタ作成
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
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

    HATUBAN_Open = False

End Function

Public Function Den_No_Set_Proc(Mode As Integer, JGYOBU As String, DEN_NO As String, Optional MSG As Integer = 1, Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      「出荷／出庫処理 入荷／入庫処理 共通」
'*          計画外伝票 伝票№発番処理
'*          大阪ＰＣ伝票№追加      2006.12.11
'*          大阪ＰＣ出庫表№追加    2007.03.15
'*
'*  計画外の伝票№の取込み
'*  (発番マスタのOPEN/CLOSEは呼び元で)
'*  引数：  モード（省略不可 10:入荷伝票№ 11:入荷テキスト№　20:出荷伝票№ 21:出荷ＩＤ№ 30:大阪PC出荷伝票№ 31:大阪PC出荷ID№ 32:大阪PC出庫表№）
'*          事業部(省略不可)
'*          伝票№(省略不可)
'*          メッセージ表示(省略可　0:表示無し　1:表示有り)
'*          リトライ(リトライ回数(0～99 0:無限))
'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer
Dim wk_No       As Long
Dim W_Cnt       As Integer
    
Dim NYU_KBN     As String * 1
Dim SYU_KBN     As String * 1

Dim NYU_ID_KBN  As String * 1
Dim SYU_ID_KBN  As String * 1

Dim OPC_ID_KBN  As String * 1
Dim OPC_DEN_KBN As String * 1


Dim c           As String * 128

    
    Den_No_Set_Proc = True
    
    DEN_NO = ""
    W_Cnt = 0
    '*------------------------------------------------------'発番マスタ読み込み
    Call UniCode_Conv(K0_HATUBAN.JGYOBU, JGYOBU)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "発番マスタ", 0)
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    If com = BtOpInsert Then
                                                            '上１桁目の区分
        If GetIni("DEN_KBN", "NYU_DEN_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_DEN_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "NYU_ID_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_ID_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_ID_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_ID_KBN = Trim(c)
        
        
        '大阪ＰＣ 追加  2006.12.11
        If GetIni("DEN_KBN", "OSAKA_ID_KBN", "SYS", c) Then
            OPC_ID_KBN = ""
        Else
            OPC_ID_KBN = Trim(c)
        End If

        If GetIni("DEN_KBN", "OSAKA_DEN_KBN", "SYS", c) Then
            OPC_DEN_KBN = ""
        Else
            OPC_DEN_KBN = Trim(c)
        End If


        
        
        
        Call UniCode_Conv(HATUBANREC.JGYOBU, JGYOBU)            '事業部
        Call UniCode_Conv(HATUBANREC.NYK_KBN, NYU_KBN)          '入荷伝票区分
        Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, "00000")       '入荷伝票№
        Call UniCode_Conv(HATUBANREC.SYK_KBN, SYU_KBN)          '出荷伝票区分
        Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, "00000")       '出荷伝票№
        
        Call UniCode_Conv(HATUBANREC.NYK_ID_KBN, NYU_ID_KBN)    '入荷ＩＤ区分
        Call UniCode_Conv(HATUBANREC.NYK_ID_NO, "00000000")     '入荷テキスト№
        Call UniCode_Conv(HATUBANREC.SYK_ID_KBN, SYU_ID_KBN)    '出荷ＩＤ区分
        Call UniCode_Conv(HATUBANREC.SYK_ID_NO, "00000000000")  '出荷ＩＤ№
        
        
        
        '大阪PC 2006.12.17
        If Trim(OPC_ID_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, OPC_ID_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "000000")
        
        End If
        
        If Trim(OPC_DEN_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, OPC_DEN_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "000000")
        
        End If
        
        
        Call UniCode_Conv(HATUBANREC.OPC_SYU_NO, "000000000000")
        
        Call UniCode_Conv(HATUBANREC.FILLER, "")
    End If
    
    Select Case Mode
        Case 10
                                    '入荷伝票№
            If StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, Format(wk_No, "00000"))
    
        Case 11
                                    '入荷ＩＤ№
            If StrConv(HATUBANREC.NYK_ID_NO, vbUnicode) = "99999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000")
            Call UniCode_Conv(HATUBANREC.NYK_ID_NO, Format(wk_No, "00000000"))
                                
        Case 20
                                '出荷伝票№
            If StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, Format(wk_No, "00000"))
        Case 21
                                    '出荷ＩＤ№
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "99999999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000000")
            Call UniCode_Conv(HATUBANREC.SYK_ID_NO, Format(wk_No, "00000000000"))
    
        Case 31
                                    '大阪ＩＤ№
            If GetIni("DEN_KBN", "SYU_ID_KBN", "SYS", c) Then
            
                SYU_ID_KBN = ""
            Else
                SYU_ID_KBN = Trim(c)
            
            End If
    
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "999999" Then
                wk_No = 1
            Else
                
                If Not IsNumeric(StrConv(HATUBANREC.OPC_ID_NO, vbUnicode)) Then
                    wk_No = 1
                Else
                    wk_No = CLng(StrConv(HATUBANREC.OPC_ID_NO, vbUnicode)) + 1
                End If
            End If
        
            
            If Trim(StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode)) = "" Then
                Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, SYU_ID_KBN)
            End If
            DEN_NO = StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode) & Format(wk_No, "000000")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, Format(wk_No, "000000"))
        Case 32
            If GetIni("DEN_KBN", "SYU_DEN_KBN", "SYS", c) Then
                SYU_KBN = ""
            Else
                SYU_KBN = Trim(c)
            End If
            
            If Trim(StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode)) = "" Then
                Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, SYU_KBN)
            End If
            
            If StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode) = "999999" Then
                wk_No = 1
            Else
                If Not IsNumeric(StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode)) Then
                    wk_No = 1
                Else
                    wk_No = CLng(StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode)) + 1
                End If
            End If
        
            DEN_NO = StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode) & Format(wk_No, "000000")
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, Format(wk_No, "000000"))
            
        Case 33
                                    '大阪出庫表№
            If StrConv(HATUBANREC.OPC_SYU_NO, vbUnicode) = "999999999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.OPC_SYU_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = Format(wk_No, "000000000000")
            Call UniCode_Conv(HATUBANREC.OPC_SYU_NO, Format(wk_No, "000000000000"))
    
    
    End Select
    '*------------------------------------------------------'発番マスタ出力
    W_Cnt = 0
    Do
        sts = BTRV(com, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, com, "発番マスタ")
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    Den_No_Set_Proc = False          '正常終了

End Function


