Attribute VB_Name = "HS_ZAIKO"
Option Explicit
'********************************************************************
'*
'*              ホスト受信データ ファイル定義
'*
'*          CREATE 2004.03.04
'********************************************************************
'ファイルＩＤ
Public Const HS_ZAIKO$ = "HS_ZAI"
'ファイル№
Public HS_Zaiko_No As Integer
'********************************************************************
'*                                                                  *
'*                           構造体定義                             *
'*                                                                  *
'********************************************************************
'*************************** 項目名定義 *****************************
'レコード定義(入庫)
Type HS_ZAIKOREC_Tag
    
    
    
    HS_JIGYOBA(0 To 7)  As Byte
    HS_HIN_GAI(0 To 19) As Byte
    HS_SHUSI(0 To 1)    As Byte
    HS_SURYO(0 To 7)    As Byte
    HS_TANA1(0 To 9)    As Byte
    HS_TANA2(0 To 9)    As Byte
    HS_TANA3(0 To 9)    As Byte
    HS_FIL(0 To 11)     As Byte
    HS_CRLF(0 To 1)     As Byte
    
    
    
'    HIN_GAI(0 To 12)    As Byte
'    HIN_TANA(0 To 9)    As Byte
'    HIN_SURYO(0 To 4)   As Byte
'    HIN_NAME1(0 To 49)  As Byte
'    HIN_NAME2(0 To 49)  As Byte
    
    
'    CR_LF(0 To 1)       As Byte           'CR.LF
    
    
    
End Type

'データ・バッファ
Public HS_ZAIKOREC As HS_ZAIKOREC_Tag
'-------------------------------------------'
Public Function HS_ZAIKO_Open(Mode As Integer, Optional JGYOBU As String = "") As Integer
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

Dim ret

    On Error GoTo HS_Zaiko_Op_Err     'ｴﾗｰﾄﾗｯﾌﾟON

    HS_ZAIKO_Open = True
                                    
    If GetIni("FILE", HS_ZAIKO, "SYS", c) Then
        Call Log_Out(LOG_F, "SYS.INI [HS_ZAIKO_SIJ]読み込みエラー")
        Exit Function
    End If
                                    
    FullPath = RTrim(c)
    
    
    If JGYOBU <> "" Then        '事業部指定有り時は事業部コードを付加する
        ret = InStr(1, Trim(FullPath), ".") - 1
        FullPath = Left(Trim(FullPath), ret) & "_" & JGYOBU & Right(Trim(FullPath), Len(Trim(FullPath)) - ret)
    End If
        
    
    
    HS_Zaiko_No = FreeFile

    Open FullPath For Input As #HS_Zaiko_No
    
    HS_ZAIKO_Open = False

    Exit Function

HS_Zaiko_Op_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
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
                ans = MsgBox("エラー [HS_ZAIKO Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
            End If
    End Select
End Function
