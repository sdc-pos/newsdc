Attribute VB_Name = "SDC_FLD_M"
Option Explicit
Public SDC_FLD_Root As String            '出力先ルート
Public SDC_FLD_Folder As String          'フォルダ名
Public SDC_FLD_File As String            'ファイル名
Public SDC_FLD_xxx As String             '識別子

Public SDC_FLD_Return As Integer         '確認画面終了状態

Function SDC_FLD_GET(Ini As String, Sect As String, Out_Nm As String) As Integer
'------------------------------------------------------------------------------
'                       出力先確認　＆　Ｇｅｔ
'
'   出力先ルートが無ければ、確認後に作成（ｷｬﾝｾﾙ時は画面ｷｬﾝｾﾙと同様）
'   フォルダが無ければ、確認後に作成（ｷｬﾝｾﾙ時は再入力）
'
'   Ini    ：iniファイル名
'   Sect   ：ｾｸｼｮﾝ名
'   Out_Nm ：出力ﾌｧｲﾙ（ﾌﾙ･ﾊﾟｽ）
'
'
'   戻り値 False：正常
'   　　　 True ：ｷｬﾝｾﾙまたはini取得異常
'------------------------------------------------------------------------------
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    '出力先ルート
    sts = GetIni(Sect, "ROOT", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_Root = Trim(c)

    'フォルダ名
    sts = GetIni(Sect, "FOLDER", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_Folder = Trim(c)

    'ファイル名
    sts = GetIni(Sect, "FILE", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_File = Trim(c)

    '識別子
    sts = GetIni(Sect, "xxx", Ini, c)
    If sts <> False Then GoTo SDC_FLD_GET_ERR
    SDC_FLD_xxx = Trim(c)

    SDC_FLD_F.Show vbModal

    If SDC_FLD_Return Then
        Out_Nm = ""
    Else
        Out_Nm = SDC_FLD_Root & "\" & SDC_FLD_Folder & _
                 "\" & SDC_FLD_File & "." & SDC_FLD_xxx
    End If

    SDC_FLD_GET = SDC_FLD_Return
    Exit Function


SDC_FLD_GET_ERR:
    SDC_FLD_GET = True
    Call Log_Out(LOG_F, Ini & " 読み込みエラー<" & Sect & ">")
    MsgBox Ini & " 読み込みエラー<" & Sect & ">", vbCritical

End Function
