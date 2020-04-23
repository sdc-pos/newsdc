Attribute VB_Name = "HS_ZAIKO_SSPC"
Option Explicit
'********************************************************************
'*
'*              ホスト受信データ ファイル定義（ＳＳＰＣ）
'*
'*          CREATE 2006.05.26
'********************************************************************
'ファイルＩＤ
Public Const HS_ZAI_SSPC$ = "HS_ZAI_SSPC"
'ファイル№
Public HS_ZAI_SSPC_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義(入庫)
Type HS_ZAI_SSPCREC_Tag

    HS_JIGYOBA_K(0 To 0)    As Byte     '事業場区分
    HS_JIGYOBA(0 To 7)      As Byte     '資産管理事業場コード
    HS_SHUSI(0 To 7)        As Byte     '在庫収支コード
    HS_FIL1(0 To 7)         As Byte
    HS_FIL2(0 To 7)         As Byte
    HS_HIN_GAI(0 To 19)     As Byte     '品目番号           '2016.03.07
'    HS_HIN_GAI(0 To 12)     As Byte     '品目番号          '2016.03.07
    HS_HIN_NAI(0 To 12)     As Byte     '工場品目番号
    HS_HIN_NAME(0 To 24)    As Byte     '品目名
    HS_TANA(0 To 7)         As Byte     'ﾛｹｰｼｮﾝ番号１
    HS_SURYO(0 To 7)        As Byte     '棚在庫数
    HS_atmark(0 To 0)       As Byte     '終端文字
    HS_CRLF(0 To 1)         As Byte     'CR + LF

End Type

'データ・バッファ
Public HS_ZAI_SSPCREC As HS_ZAI_SSPCREC_Tag
'-------------------------------------------'
Public Function HS_ZAI_SSPC_Open(Mode As Integer, Optional JGYOBU As String = "") As Integer
'********************************************************************
'*
'*      ホスト受信データ  ＯＰＥＮ
'*
'*      引数　:OPENモード（0:参照　1:更新）
'*             ﾃﾞｰﾀﾀｲﾌﾟ   (1:入庫　2:出荷)
'*
'*      戻り値:false 正常
'*             true  異常
'*
'*          CREATE 2004.03.05
'********************************************************************

Dim ans         As Integer
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret

    On Error GoTo HS_ZAI_SSPC_Op_Err     'ｴﾗｰﾄﾗｯﾌﾟON

    HS_ZAI_SSPC_Open = True
                                    
    If GetIni("FILE", HS_ZAI_SSPC, "SYS", c) Then
        Call LOG_OUT(LOG_F, "SYS.INI [HS_ZAI_SSPC_SIJ]読み込みエラー")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    
    If JGYOBU <> "" Then        '事業部指定有り時は事業部コードを付加する
        Ret = InStr(1, Trim(FullPath), ".") - 1
        FullPath = Left(Trim(FullPath), Ret) & "_" & JGYOBU & Right(Trim(FullPath), Len(Trim(FullPath)) - Ret)
    End If
        
    
    
    HS_ZAI_SSPC_No = FreeFile

    Open FullPath For Binary As #HS_ZAI_SSPC_No
    
    HS_ZAI_SSPC_Open = False

    Exit Function

HS_ZAI_SSPC_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
                If ans = vbYes Then
                    Resume
                End If
            End If
        Case ErrDeviceUnavailable
            If Mode = 1 Then
                Beep
                ans = MsgBox("ドライブまたはパスが見つかりません" & FullPath, vbExclamation)
            End If
        Case ErrNotFound
            If Mode = 1 Then
                Beep
                ans = MsgBox("ファイルが見つかりません" & FullPath, vbExclamation)
            End If
        Case Else
            If Mode = 1 Then
                Beep
                ans = MsgBox("エラー [HS_ZAI_SSPC Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
