Attribute VB_Name = "Tool"
Option Explicit
'****************************************************
'*      グローバル定義
'*
'*
'****************************************************
Type JGYOBU_TBL                 '有効事業部テーブル
    CODE As String * 1
    NAME As String * 20
    COLOR As Long
End Type
Public JGYOBU_T() As JGYOBU_TBL

Public Last_JGYOBU As String    '現在処理中事業部コード

Public LOG_F  As String         'ログファイル名称




'[ 数値内容チェック，編集処理 ]用定数
'                           ＜　処理タイプ　＞
Public Const CHK_EDIT% = 0              'チェック＆編集
Public Const EDIT_ONLY% = 1             '編集のみ
'                           ＜　負数可，不可　＞
Public Const NEGA_DIS% = 0              '不可
Public Const NEGA_ENA% = 1              '可
'                           ＜　ゼロ抑制　＞
Public Const ZSUP_DIS% = 0              '無し
Public Const ZSUP_ENA% = 1              '有り
'                           ＜　カンマ編集　＞
Public Const COMA_ENA% = 0              '有り
Public Const COMA_DIS% = 1              '無し
'[ カレントフォームハードコピー ]用定数
'                           ＜　キーコード定数　＞
Public Const VK_LMENU = &HA4                'Altキー
Public Const VK_SNAPSHOT = &H2C             'PrintScreenキー
'                           ＜　ｷｰﾎﾞｰﾄﾞｲﾍﾞﾝﾄﾌﾗｸﾞ定数　＞
Public Const KEYEVENTF_EXTENDEDKEY = &H1    'キーを押す
Public Const KEYEVENTF_KEYUP = &H2          'キーを離す

'[ システム予約済要因コード ]用項目
'Public YOIN_HENPIN          As String * 2       '「良品返品」の要因
'2004Public YOIN_MAE_SOUSAI      As String * 2       '「前借り相殺」の要因
'Public YOIN_SIKYU           As String * 2       '「支給」の要因
'2004Public YOIN_CHOKUSO         As String * 2       '「出庫(直送)」の要因
'2004Public YOIN_CHOKU_MODOSI    As String * 2       '「出庫(直送)の戻し」の要因
'2004Public YOIN_HSP             As String * 2       '「出荷（補／ス）」の要因
'2004Public YOIN_TUK             As String * 2       '「出荷（月切）」の要因
'2004Public YOIN_SPO             As String * 2       '「出荷（スポット）」の要因
'2004Public YOIN_HJU             As String * 2       '「出荷（補充）」の要因
'2004Public YOIN_TOK             As String * 2       '「出荷（特売）」の要因
'2004Public YOIN_BOU             As String * 2       '「出荷（貿易）」の要因
'2004Public YOIN_SYU_HSP         As String * 2       '「出荷（補／ス）出庫表出庫」の要因
'2004Public YOIN_SYU_TUK         As String * 2       '「出荷（月切）出庫表出庫」の要因
'2004Public YOIN_SYU_SPO         As String * 2       '「出荷（スポット）出庫表出庫」の要因
'2004Public YOIN_SYU_HJU         As String * 2       '「出荷（補充）出庫表出庫」の要因
'2004Public YOIN_SYU_TOK         As String * 2       '「出荷（特売）出庫表出庫」の要因
'2004Public YOIN_SYU_BOU         As String * 2       '「出荷（貿易）出庫表出庫」の要因
'2004Public YOIN_KIN             As String * 2       '「出荷（緊急）」の要因
'Public YOIN_NYUKA           As String * 2       '「通常入庫（入荷倉庫より）」の要因


Sub File_Error(sts As Integer, Opretion As Integer, file As String, Optional Mode As Integer = 1)
'****************************************************
'*      ファイルエラー処理
'*
'*  引  数: ファイルステータス
'*          オペレーションコード
'*          ファイル名称
'*          モード 1: 表示有り 0: 表示無し
'*
'*  戻り値: なし
'*          CREATE 1997.01.09  M.Yoshizawa                          *
'****************************************************
    Dim Buf As String
    Buf = "Op= " + Str$(Opretion) + " " + "sts = " + Str$(sts) + " " + file
    Call Log_Out(LOG_F, Buf)
    
    If Mode = 1 Then
        Call Bt_Error(sts, Opretion, file)
    End If
End Sub
Sub Ctrl_Lock(F_Obj As Form)
'*****************************************************
'*　　　コントロール　ロック
'*
'*　引　数：フォームオブジェクト
'*
'*　戻り値：なし
'*          CREATE 1999.03.16  S.Shibano
'*****************************************************
Dim i As Integer

    For i = 0 To F_Obj.Count - 1
                                    '「Enabled」を持つｵﾌﾞｼﾞｪｸﾄか？
        If TypeOf F_Obj.Controls(i) Is CommandButton Or _
           TypeOf F_Obj.Controls(i) Is ComboBox Or _
           TypeOf F_Obj.Controls(i) Is CheckBox Or _
           TypeOf F_Obj.Controls(i) Is DirListBox Or _
           TypeOf F_Obj.Controls(i) Is TextBox Or _
           TypeOf F_Obj.Controls(i) Is DriveListBox Or _
           TypeOf F_Obj.Controls(i) Is FileListBox Or _
           TypeOf F_Obj.Controls(i) Is ListBox Or _
           TypeOf F_Obj.Controls(i) Is HScrollBar Or _
           TypeOf F_Obj.Controls(i) Is VScrollBar Then
        
        
        
        
            F_Obj.Controls(i).Tag = F_Obj.Controls(i).Enabled
            F_Obj.Controls(i).Enabled = False
        End If
    
    
    Next i

End Sub

Sub Ctrl_UnLock(F_Obj As Form)
'*****************************************************
'*　　　コントロール　アンロック
'*
'*　引　数：フォームオブジェクト
'*
'*　戻り値：なし
'*          CREATE 1999.03.16  S.Shibano
'*****************************************************
Dim i As Integer

    For i = 0 To F_Obj.Count - 1
                                    '「Enabled」を持つｵﾌﾞｼﾞｪｸﾄか？
        If TypeOf F_Obj.Controls(i) Is CommandButton Or _
           TypeOf F_Obj.Controls(i) Is ComboBox Or _
           TypeOf F_Obj.Controls(i) Is CheckBox Or _
           TypeOf F_Obj.Controls(i) Is DirListBox Or _
           TypeOf F_Obj.Controls(i) Is TextBox Or _
           TypeOf F_Obj.Controls(i) Is DriveListBox Or _
           TypeOf F_Obj.Controls(i) Is FileListBox Or _
           TypeOf F_Obj.Controls(i) Is ListBox Or _
           TypeOf F_Obj.Controls(i) Is HScrollBar Or _
           TypeOf F_Obj.Controls(i) Is VScrollBar Then
        
           F_Obj.Controls(i).Enabled = F_Obj.Controls(i).Tag
        End If
    Next i


End Sub

Function GetIni(Section As String, ITEM As String, NAME As String, c As String) As Integer
'****************************************************
'*      ＩＮＩファイル取り込み処理
'*
'*  引  数: セクション名
'*          アイテム名
'*          ＩＮＩファイル名
'*          取り込み領域（ＯＵＴＰＵＴ）
'*
'*  戻り値: false 正常
'*          true  異常
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim fileName        As String
Dim sts             As Long
Dim Work(0 To 127)  As Byte
Dim buf1            As String * 128
Dim buf2            As String
    
    GetIni = False
    fileName = App.Path
    If Right(fileName, 1) <> "\" Then
        fileName = fileName & "\"
    End If
    fileName = fileName & NAME & ".ini"
    c = Space(Len(c))
    sts = GetPrivateProfileString(Section, ITEM, "", buf1, 128, fileName)
    If sts = False Then
        GetIni = True
    Else
        buf2 = RTrim(buf1)
        Call UniCode_Conv(Work, buf2)
        c = StrConv(LeftB(Work, sts), vbUnicode)
    End If
End Function
Function WriteIni(Section As String, ITEM As String, NAME As String, c As String) As Integer
'****************************************************
'*      ＩＮＩファイル書き込み処理
'*
'*  引  数: セクション名
'*          アイテム名
'*          ＩＮＩファイル名
'*          書き込み内容
'*
'*  戻り値: false 正常
'*          true  異常
'*          CREATE 1997.02.15  M.Yoshizawa
'****************************************************
Dim fileName As String
Dim sts As Long
    
    WriteIni = False
    fileName = App.Path
    If Right(fileName, 1) <> "\" Then
        fileName = fileName & "\"
    End If
    fileName = fileName & NAME & ".ini"
    sts = WritePrivateProfileString(Section, ITEM, c, fileName)
    If sts = False Then
        WriteIni = True
    End If

End Function


Sub Log_Out(file As String, MSG As String)
'****************************************************
'*      ログファイル出力処理
'*
'*  引  数: ログファイル名
'*          出力内容
'*
'*  戻り値: なし
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

    
    stream = FreeFile
    Open file For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "???"
    End If
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog & " " & MSG)
    Print #stream, Buf
    Close stream
End Sub

Sub UniCode_Conv(Buffer() As Byte, Unicode As String)
'****************************************************
'*      ＵＮＩＣＯＤＥ変換
'*
'*  引  数: ＡＮＳＩ（ＯＵＴＰＵＴ）
'*          ＵＮＩＣＯＤＥ
'*
'*  戻り値: なし
'*          CREATE 1997.01.09  M.Yoshizawa
'****************************************************
Dim TmpBuf() As Byte
Dim TmpStr As String
Dim TmpStrlen As Integer
Dim i As Integer
Dim Swork As String
                            '初期化
    Swork = Space(UBound(Buffer) + 1)
    TmpBuf = ""
    TmpStr = StrConv(Swork, vbFromUnicode)
    TmpStrlen = LenB(TmpStr) - 1
    TmpBuf = StrConv(Swork, vbFromUnicode)
    For i = 0 To TmpStrlen
        Buffer(i) = TmpBuf(i)
    Next i

                            '変換
    TmpBuf = ""
    TmpStr = StrConv(Unicode, vbFromUnicode)
    TmpStrlen = LenB(TmpStr) - 1
    TmpBuf = StrConv(Unicode, vbFromUnicode)
    For i = 0 To TmpStrlen
                            '受け取り側の桁数を超えた場合は切り捨てす
        If i > (UBound(Buffer)) Then
           Exit For
        End If
        
        Buffer(i) = TmpBuf(i)
    Next i
End Sub



Function Numeric_Check(Mode As Integer, Keta As Integer, Dec As Integer, NEGA As Integer, ZSUP As Integer, COMA As Integer, Buf As String, RetBuf As String) As Integer
'*****************************************************
'*　　　数値内容チェック，編集処理
'*
'*　引　数：処理タイプ（０：チェック＆編集
'*　　　　　　　　　　　１：編集のみ）
'*　　　　　桁数（小数点，符号，カンマ含む）
'*　　　　　小数桁数
'*　　　　　負数可，不可（０：不可，１：可）
'*　　　　　ゼロ抑制　　（０：不可，１：可）
'*　　　　　カンマ編集　（０：有り，１：無し）
'*　　　　　チェック対象
'*　　　　　編集内容
'*
'*　戻り値：ｆａｌｓｅ　正常
'*　　　　　ｔｒｕｅ　　異常
'*          CREATE 1997.01.09  M.Yoshizawa
'*****************************************************
Dim Using_Value As String
Dim Using_wk As String
Dim dNum As Double
Dim iLen As Integer
Dim iSei_Len As Integer
Dim iDec_Len As Integer
Dim iDec_Pos As Integer
Dim iGW_EDIT_pos As Integer
Dim iKeta_cnt As Integer
Dim GW_EDIT_Str As String
    
On Error GoTo Error_Proc
    
    Numeric_Check = True
    RetBuf = Space(Keta)
    Using_wk = Trim(Buf)
    
    'パラメータチェック
    If Mode <> CHK_EDIT And Mode <> EDIT_ONLY Then Exit Function
    If Keta < 0 Or Dec < 0 Then Exit Function
    If NEGA <> NEGA_DIS And NEGA <> NEGA_ENA Then Exit Function
    If ZSUP <> ZSUP_DIS And ZSUP <> ZSUP_ENA Then Exit Function
    If COMA <> COMA_ENA And COMA <> COMA_DIS Then Exit Function
    
    If (IsNumeric(Using_wk) = False) Then   '数値以外エラー
        Exit Function
    End If
    
    dNum = CDbl(Using_wk)
    iDec_Pos = InStr(Using_wk, ".")         '小数点の位置（０＝無し）
    If iDec_Pos = 0 Then
        iDec_Len = 0
    Else
        iDec_Len = Len(Mid(Using_wk, iDec_Pos + 1)) '小数点以下の桁数
    End If
    
    If Mode = EDIT_ONLY Then GoTo Numeric_EDIT      '*** ->ﾁｪｯｸ ｽｷｯﾌﾟ ***
    
    If NEGA = NEGA_DIS And (Sgn(dNum) < 0) Then    'マイナス不可でマイナス値
        Exit Function
    End If

    If Dec < iDec_Len Then                  '小数点以下の桁数オーバー
        Exit Function
    End If
    
Numeric_EDIT:       '*** 編集フォーマット作成 ***
    
                        '** 編集後の整数部桁数チェック **
    If Keta = 0 Then        '桁数無指定
        Using_Value = "#0"
    Else                    '桁数指定
        If Dec = 0 Then             '小数点無し
            iSei_Len = Keta
        Else                        '小数点有り
            iSei_Len = Keta - Dec - 1
        End If
        If iSei_Len <= 0 Then Exit Function     '整数部桁数不足エラー
                    '*** 編集文字列作成 ***
        If COMA = COMA_ENA Then                  'カンマ有り
            If ZSUP = ZSUP_DIS Then                  'ゼロサプレス無し
                GW_EDIT_Str = "0"
                If NEGA = NEGA_ENA Then
                    iSei_Len = iSei_Len - 1     'マイナス可なら1桁減らす
                End If
            Else                            'ゼロサプレス
                GW_EDIT_Str = "#"
            End If
            Using_Value = "0"
            iKeta_cnt = 1
            For iGW_EDIT_pos = 1 To iSei_Len - 1
                If (iKeta_cnt Mod 3) = 0 Then
                    iGW_EDIT_pos = iGW_EDIT_pos + 1
                    If iGW_EDIT_pos < iSei_Len Then
                        Using_Value = GW_EDIT_Str & "," & Using_Value
                    End If
                Else
                    Using_Value = GW_EDIT_Str & Using_Value
                End If
                iKeta_cnt = iKeta_cnt + 1
            Next iGW_EDIT_pos
        Else                            'カンマ無し
            If ZSUP = ZSUP_DIS Then          'ゼロサプレス無し
                If Sgn(dNum) < 0 Then
                    Using_Value = String(iSei_Len - 1, "0") '値がマイナスなら1桁減らす
                Else
                    Using_Value = String(iSei_Len, "0")
                End If
            Else                            'ゼロサプレス
                Using_Value = String(iSei_Len - 1, "#") & "0"
            End If
        End If
    End If

    If Dec > 0 Then                 '小数点以下
        Using_Value = Using_Value & "." & String(Dec, "0")
    End If
    
    iLen = Len(Using_Value)
    If Keta = 0 Then        '桁数無指定
        RetBuf = Format(dNum, Using_Value)
    Else                    '桁数指定
        If ZSUP = ZSUP_DIS Then      'ゼロサプレス無しで〜
            'カンマ有り & マイナス可 か？
            'カンマ無し & マイナス値 なら1桁増やす
            If (COMA = COMA_ENA And NEGA = NEGA_ENA) Or _
               (COMA = COMA_DIS And Sgn(dNum) < 0) Then
                iLen = iLen + 1
            End If
        End If
        If iLen <> Keta Then Exit Function      '->編集桁数不一致
        Using_wk = Format(dNum, Using_Value)
        iLen = Len(Using_wk)
        Select Case iLen            '編集後桁数
          Case Keta
            RetBuf = Using_wk
          Case Is < Keta
            RetBuf = Space(Keta - iLen) & Using_wk
          Case Else                     '桁数オーバー
            Exit Function
        End Select
    End If
    
    Numeric_Check = False
    
Exit Function

Error_Proc:

    Numeric_Check = True

End Function
Function JGYOB_TB_Set(Optional JGYOBU As Integer = 0) As Integer
'****************************************************
'*      事業部テーブルセット
'*
'*  戻り値: false 正常
'*          true  異常
'*          CREATE 1997.07.05  S.Shibano
'****************************************************
Dim c   As String
Dim i   As Long
Dim j   As Integer

    JGYOB_TB_Set = False

'    For i = 0 To UBound(JGYOBU_T)
'        JGYOBU_T(i).Code = " "
'        JGYOBU_T(i).NAME = "                    "
'    Next i

                                '事業部取り込み
    i = 0
    j = 0
    Do
        If GetIni("JIGYOBU", "code" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Call Log_Out(LOG_F, "[SYS.INI] [JIGYOBU] [CODE] READ ERROR")
            JGYOB_TB_Set = True
            Exit Function
        End If
        If RTrim(c) = "0" Then
            Exit Do
        End If

        If JGYOBU = 1 And _
            RTrim(c) = SHIZAI Then
            '資材を無視
        Else
            ReDim Preserve JGYOBU_T(j)
    
            JGYOBU_T(j).CODE = RTrim(c)
            If GetIni("JIGYOBU", "name" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
                Call Log_Out(LOG_F, "[SYS.INI] [JIGYOBU] [NAME] READ ERROR")
                JGYOB_TB_Set = True
                Exit Function
            End If
            JGYOBU_T(j).NAME = RTrim(c)
    
            If GetIni("JIGYOBU", "color" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
                Call Log_Out(LOG_F, "[SYS.INI] [JIGYOBU] [COLOR] READ ERROR")
                JGYOB_TB_Set = True
                Exit Function
            End If
            JGYOBU_T(j).COLOR = CLng(RTrim(c))
            j = j + 1
        End If
        i = i + 1
    
    Loop
                                    
                                'デフォルト事業部取り込み
    If GetIni("JIGYOBU", "DEF_NO", "SYS", c) Then
         Call Log_Out(LOG_F, "[SYS.INI] [JIGYOBU] [DEF_NO] READ ERROR")
        JGYOB_TB_Set = True
        Exit Function
    End If
    Last_JGYOBU = RTrim(c)

End Function

Public Sub Data_Select(In_Dat As String, Get_Pos As Integer, Max_Pos As Integer, Out_Dat As String)
'****************************************************
'*      データの切り出し
'*　引　数：切り出し元データ(","区切りのデータ)
'*　　　　　切り出しポジション
'*　　　　　最大個数
'*　　　　　切り出されたデータ
'*
'*  戻り値: なし
'*          CREATE 2001.04.10  M.Yoshizawa
'****************************************************

Dim i           As Integer
Dim Start_Pos   As Integer
Dim End_Pos     As Integer

    Out_Dat = ""

    Start_Pos = 1
    For i = 1 To Max_Pos
        End_Pos = InStr(Start_Pos, In_Dat, ",")
        If End_Pos = 0 And i <> Max_Pos Then
            Exit Sub
        End If
    
        If Get_Pos = i Then
            If End_Pos > Start_Pos Then
                Out_Dat = Mid(In_Dat, Start_Pos, End_Pos - Start_Pos)
            Else
                Out_Dat = Mid(In_Dat, Start_Pos)
            End If
            If Out_Dat = "NON" Then
                Out_Dat = ""
            End If
            Exit Sub
        End If
        Start_Pos = End_Pos + 1
    Next i

End Sub

'Public Function SYSTEM_YOIN_Set() As Integer
''****************************************************
''*      システム予約済要因の取込み
''*
''*  引数 :  なし
''*  戻り値: false       正常
''*          SYS_ERR     継続できない異常
''****************************************************
'Dim c As String
'
'    SYSTEM_YOIN_Set = SYS_ERR
'
'
'
'                                        '「通常入荷」の要因
'    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TU_NYUKA] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TU_NYUKA = Trim(c)
'                                        '「前借り入荷」の要因
'    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_MAEGARI = Trim(c)
'                                        '「良品返品」の要因
''    If GetIni("YOIN", "YOIN_HENPIN", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HENPIN] READ ERROR")
''        Exit Function
''    End If
''    YOIN_HENPIN = Trim(c)
'                                        '「前借り相殺」の要因
'    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAE_SOUSAI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_MAE_SOUSAI = Trim(c)
'                                        '「支給」の要因
''    If GetIni("YOIN", "YOIN_SIKYU", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SIKYU] READ ERROR")
''        Exit Function
''    End If
''   YOIN_SIKYU = Trim(c)
'                                        '「出庫(直送)」の要因
'    If GetIni("YOIN", "YOIN_CHOKUSO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_CHOKUSO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_CHOKUSO = Trim(c)
'                                        '「出庫(直送)戻し」の要因
'    If GetIni("YOIN", "YOIN_CHOKU_MODOSI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_CHOKU_MODOSI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_CHOKU_MODOSI = Trim(c)
'                                        '「出荷（補／ス）」の要因
'    If GetIni("YOIN", "YOIN_HSP", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HSP] READ ERROR")
'        Exit Function
'    End If
'    YOIN_HSP = Trim(c)
'                                        '「出荷（月切）」の要因
'    If GetIni("YOIN", "YOIN_TUK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TUK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TUK = Trim(c)
'                                        '「出荷（スポット）」の要因
'    If GetIni("YOIN", "YOIN_SPO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SPO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SPO = Trim(c)
'                                        '「出荷（補充）」の要因
'    If GetIni("YOIN", "YOIN_HJU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_HJU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_HJU = Trim(c)
'                                        '「出荷（特売）」の要因
'    If GetIni("YOIN", "YOIN_TOK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TOK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_TOK = Trim(c)
'                                        '「出荷（貿易）」の要因
'    If GetIni("YOIN", "YOIN_BOU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_BOU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_BOU = Trim(c)
'                                        '「出荷（補／ス）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_HSP", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_HSP] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_HSP = Trim(c)
'                                        '「出荷（月切）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_TUK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_TUK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_TUK = Trim(c)
'                                        '「出荷（スポット）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_SPO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_SPO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_SPO = Trim(c)
'                                        '「出荷（補充）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_HJU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_HJU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_HJU = Trim(c)
'                                        '「出荷（特売）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_TOK", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_TOK] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_TOK = Trim(c)
'                                        '「出荷（貿易）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_SYU_BOU", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_SYU_BOU] READ ERROR")
'        Exit Function
'    End If
'    YOIN_SYU_BOU = Trim(c)
'                                        '「出荷（緊急）出庫表出庫」の要因
'    If GetIni("YOIN", "YOIN_KIN", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_KIN] READ ERROR")
'        Exit Function
'    End If
'    YOIN_KIN = Trim(c)
'                                        '「通常入庫（入荷倉庫より）」の要因
''    If GetIni("YOIN", "YOIN_NYUKA", "SYS", c) Then
''        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_NYUKA] READ ERROR")
''        Exit Function
''    End If
''    YOIN_NYUKA = Trim(c)
'                                        '「国内外振替え」の要因
'    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE = Trim(c)
'                                        '「国内外振替え事の出庫」の要因
'    If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_OUT] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE_OUT = Trim(c)
'                                        '「国内外振替え事の入庫」の要因
'    If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_IN] READ ERROR")
'        Exit Function
'    End If
'    YOIN_FURIKAE_IN = Trim(c)
'
'                                        '「WEL 棚卸し」の要因
'    If GetIni("YOIN", "YOIN_WEL_TANAOROSI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANAOROSI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANAOROSI = Trim(c)
'                                        '「WEL 棚番表示」の要因
'    If GetIni("YOIN", "YOIN_WEL_TANAHYOJI", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANAHYOJI] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANAHYOJI = Trim(c)
'                                        '「WEL 棚照合」の要因
'    If GetIni("YOIN", "YOIN_WEL_TANASHOGO", "SYS", c) Then
'        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANASHOGO] READ ERROR")
'        Exit Function
'    End If
'    YOIN_WEL_TANASHOGO = Trim(c)
'
'
'    SYSTEM_YOIN_Set = False
'End Function


Sub Tab_Ctrl(Sf As Integer)
'******************************************************
'*　　　タブコントロール
'*
'*　引　数：Shift  (Shiftのみ)
'*
'*　戻り値：なし
'******************************************************
Dim S_Wk As String

    S_Wk = ""
    If Sf = vbShiftMask Then S_Wk = "+"
    S_Wk = S_Wk & "{TAB}"
    SendKeys S_Wk           ', True

End Sub

Sub Form_HCopy(obj_Pic As Object, pr_Size As Integer, pr_Orient As Integer)
'00/02/12「ＨＭ」より引用
'---------------------------------------------------------------------------
'           カレントフォームのハードコピー
'
'［引数］obj_Pic   ：ｲﾒｰｼﾞ取込み用ﾋﾟｸﾁｬｰｵﾌﾞｼﾞｪｸﾄ（FORMの見えない位置に配置）
'　　　　pr_Size   ：印刷用紙サイズ
'　　　　pr_Orient ：印刷用紙方向
'
'《キー操作について》
'　　Ｗｉｎ９５／９８ではキーを「押す」「離す」をまとめて行えるが、
'　　ＷｉｎＮＴでは「押す」「離す」を別々にしないと認識してくれない
'
'《ハードコピー使用上の注意》
'　サブCALL時点でフォーカスを持つFORMが印刷される。
'　一旦クリップボードに取り込んだ画像を、ピクチャボックスに読み込んで印刷する為、
'　画像読み込み用のピクチャボックスコントロールを引数として渡す。
'　ピクチャボックスは、FORM上の見えない位置に配置するか、Visible=Falseにする。
'
'---------------------------------------------------------------------------
Dim sngPrnRatio As Single
Dim sngPrnHeight As Single
Dim sngPrnWidth As Single
Dim sngPicPosX As Single
Dim sngPicPosY As Single
Dim sngPicRatio As Single
Dim sngPicWidth As Single
Dim sngPicHeight As Single

Dim c As String
Dim USE_Printer As String
Dim Wk_Printer As Printer

Dim Pri_Name As Printer





'ハードコピー用プリンタを選択（システムプリンタ）
'''    If GetIni("SYSTEM", "PRINTER", "SYS", c) Then
'''        Beep
'''        MsgBox "システムプリンタが定義されていません。", vbCritical
'''        Exit Sub
'''   End If
'''    USE_Printer = RTrim(c)


    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            USE_Printer = Pri_Name.DeviceName
            Exit For
        End If
    Next


    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If Wk_Printer.DeviceName = USE_Printer Then
            Set Printer = Wk_Printer
            Exit For
        End If
    Next

'クリップボードをクリア
    Clipboard.Clear

'Altキーを押す
    Keybd_Event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
'PrintScreenキーを押す
    Keybd_Event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
'キー操作を実行（重要：これが無いとﾌﾟﾛｼｰｼﾞｬを抜ける迄キー操作が発生しない）
    DoEvents
'PrintScreenキーを離す
    Keybd_Event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
'Altキーを離す
    Keybd_Event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
'クリップボードからフォームの画像を取得
    obj_Pic.Picture = Clipboard.GetData()
'画像の印刷位置とサイズを修正
    With obj_Pic.Picture
        sngPicRatio = .Width / .Height
    End With

    With Printer
        '印刷用紙の設定
        .PaperSize = pr_Size         '用紙サイズ
        .Orientation = pr_Orient     '上にして印刷する用紙の辺（長辺，短辺）

        '印刷用紙の設定
        sngPrnRatio = .ScaleWidth / .ScaleHeight
        sngPrnWidth = .ScaleX(.ScaleWidth, _
                              .ScaleMode, _
                              vbHimetric)
        sngPrnHeight = .ScaleY(.ScaleHeight, _
                               .ScaleMode, _
                               vbHimetric)
        If sngPicRatio > sngPrnRatio Then
            sngPicHeight = _
                .ScaleY(sngPrnWidth / sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnWidth, _
                        vbHimetric, _
                        .ScaleMode)
        Else
            sngPicHeight = _
                .ScaleY(sngPrnHeight, _
                        vbHimetric, _
                        .ScaleMode)
            sngPicWidth = _
                .ScaleX(sngPrnHeight * sngPicRatio, _
                        vbHimetric, _
                        .ScaleMode)
        End If
        sngPicPosX = (.ScaleWidth - sngPicWidth) / 2
        sngPicPosY = (.ScaleHeight - sngPicHeight) / 2

        'フォームの画像を印刷
        .PaintPicture obj_Pic.Picture, _
                      sngPicPosX, _
                      sngPicPosY, _
                      sngPicWidth, _
                      sngPicHeight
        '印刷を終了し、制御をプリンタに渡す
        .EndDoc
    End With
End Sub


