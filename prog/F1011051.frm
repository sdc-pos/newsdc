VERSION 5.00
Begin VB.Form F1011051 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "棚番別個装箱･ランク自動生成"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4575
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3900
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "生成データ(累計)"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   6
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label L_CNT 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label L_CNT 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "処理中ランク"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   3
      Top             =   1620
      Width           =   1440
   End
   Begin VB.Label L_CNT 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "処理中個装箱№"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "棚別個装箱マスタデータ生成中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   3990
   End
End
Attribute VB_Name = "F1011051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const plabPacking_No% = 0       '個装箱№
Private Const plabRank% = 1             'ランク
Private Const plaData_Create% = 2       'データ生成件数

Private Const RANK_CHR$ = "A-1,A-2,B-1,B-2,C-1,C-2,D,E"
Dim T_Rank              As Variant      'ランクテーブル

Dim T_SokoChr(26)       As String       '倉庫文字⇒倉庫№ 読替ﾃｰﾌﾞﾙ
Dim wPUT_CNT            As Long         'データ生成件数(表示用ｶｳﾝﾀ)

Private Sub Command_Click(Index As Integer)

    If Data_Create_Proc Then
        MsgBox "異常発生の為、強制終了されました。", vbOKOnly
    End If

    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim yn      As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

    'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If


    '実行確認
    yn = vbYes
    Beep
    yn = MsgBox("棚別個装箱マスタデータ生成、実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If yn = vbYes Then
        Show
        DoEvents
        Command(0).Value = True     '処理開始
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If

    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    Set F1011051 = Nothing

    End

End Sub

Private Function Data_Create_Proc() As Integer
'========================================================================================
'                   棚別個装箱マスタデータ生成処理
'========================================================================================
Dim i           As Integer
Dim j           As Integer
Dim wCSV        As Variant
Dim wStr        As String
Dim FNo         As Integer
Dim c           As String * 128
Dim sts         As Integer


    Data_Create_Proc = True


    '棚別個装箱マスタＯＰＥＮ
    If TPACKING_ReCreate Then           'ﾌｧｲﾙ再作成
        Exit Function
    End If

    If TPACKING_Open(BtOpenNomal) Then  'OPEN
        Exit Function
    End If


    '倉庫番号読替え文字の取得(iniﾌｧｲﾙより)
    For i = 1 To UBound(T_SokoChr)
        If GetIni("SOKO_NO", Format(i, "00"), "SYS", c) Then
        Else
            T_SokoChr(i) = RTrim(c)
        End If
    Next i


    'ランクテーブル初期化
    T_Rank = Split(RANK_CHR, ",", -1, vbTextCompare)


    '棚別個装箱マスタ設定用ＣＳＶ　ＯＰＥＮ
    If GetIni("FILE", "TPACKING_CSV", "SYS", c) Then
        Beep
        MsgBox "棚別個装箱マスタＣＳＶファイル名の獲得に失敗しました。"
        GoTo Data_Create_Proc_Exit
    End If
    wStr = RTrim(c)

    FNo = FreeFile
    Open wStr For Input As #FNo

    '棚別個装箱マスタ設定用ＣＳＶ　読込み
    L_CNT(plaData_Create).Caption = ""              'データ生成件数 クリア
    wPUT_CNT = 0

    Do While EOF(FNo) = False
        Line Input #FNo, wStr
        wCSV = Split(wStr, ",", -1, vbTextCompare)

        If Trim(wCSV(0)) = "No" Then
        Else
'        If Left(wCSV(0), 1) = "D" Then
            
            
            For i = 0 To UBound(T_Rank)
                j = i * 2 + 1
                If wCSV(j) <> "" And wCSV(j + 1) <> "" Then
                    '棚別個装箱マスタデータ書込み
                    If Data_Put_Proc(CStr(wCSV(0)), CStr(wCSV(j)), _
                                     CStr(wCSV(j + 1)), CStr(T_Rank(i))) Then
                        Close FNo
                        GoTo Data_Create_Proc_Exit
                    End If
                End If
            Next i
'        End If
        End If
    Loop

    '棚別個装箱マスタ設定用ＣＳＶ　ＣＬＯＳＥ
    Close FNo


    Label1(0).Caption = ""
    DoEvents

    Data_Create_Proc = False
    Beep
    MsgBox "データ生成処理が正常に終了しました。"



Data_Create_Proc_Exit:

    '棚別個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚別個装箱マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If


End Function

Private Function Data_Put_Proc(pPACKNo As String, pSRANGE As String, pERANGE As String, pRANK As String) As Integer
'========================================================================================
'                   棚別個装箱マスタデータ書込み
'========================================================================================
Dim sts         As Integer
Dim com         As Integer

Dim wSoko       As String       '倉庫番号
Dim wRetu_S     As String       '列 開始
Dim wRetu_E     As String       '列 終了
Dim wRen_S      As String       '連 開始
Dim wRen_E      As String       '連 終了
Dim i           As Integer



    Data_Put_Proc = True


    '進捗表示
    L_CNT(plabPacking_No).Caption = pPACKNo         '処理中 個装箱№
    L_CNT(plabRank).Caption = pRANK                 '処理中 ランク


    '倉庫番号チェック
    wSoko = ""
    For i = 1 To UBound(T_SokoChr)
        If T_SokoChr(i) = Left(pSRANGE, 1) Then
            wSoko = Format(i, "00")
            Exit For
        End If
    Next i

    If wSoko = "" Then
        Beep
        MsgBox "指定文字に該当する倉庫が見つかりません" & vbCrLf & vbCrLf & _
               "個装箱№：" & pPACKNo & vbCrLf & _
               "ランク　：" & pRANK & vbCrLf & _
               "範囲指定：" & pSRANGE & "～" & pERANGE
        Exit Function
    End If

    If Left(pSRANGE, 1) <> Left(pERANGE, 1) Then
        Beep
        MsgBox "倉庫を跨ぐ範囲は指定できません" & vbCrLf & vbCrLf & _
               "個装箱№：" & pPACKNo & vbCrLf & _
               "ランク　：" & pRANK & vbCrLf & _
               "範囲指定：" & pSRANGE & "～" & pERANGE
        Exit Function
    End If


    '範囲開始～範囲終了チェック
    If InStr(pSRANGE, "-") > 0 And InStr(pERANGE, "-") > 0 Then
    Else
        Beep
        MsgBox "範囲指定が正しくありません" & vbCrLf & vbCrLf & _
               "個装箱№：" & pPACKNo & vbCrLf & _
               "ランク　：" & pRANK & vbCrLf & _
               "範囲指定：" & pSRANGE & "～" & pERANGE
        Exit Function
    End If

    i = InStr(pSRANGE, "-")
    wRetu_S = Mid(pSRANGE, 2, i - 2)    '列 開始
    wRen_S = Mid(pSRANGE, i + 1, 2)     '連 開始

    i = InStr(pERANGE, "-")
    wRetu_E = Mid(pERANGE, 2, i - 2)    '列 終了
    wRen_E = Mid(pERANGE, i + 1, 2)     '連 終了

    If IsNumeric(wRetu_S) And IsNumeric(wRen_S) And _
       IsNumeric(wRetu_E) And IsNumeric(wRen_E) Then
    Else
        Beep
        MsgBox "範囲指定が正しくありません" & vbCrLf & vbCrLf & _
               "個装箱№：" & pPACKNo & vbCrLf & _
               "ランク　：" & pRANK & vbCrLf & _
               "範囲指定：" & pSRANGE & "～" & pERANGE
        Exit Function
    End If

    wRetu_S = Format(Val(wRetu_S), "00")    '列 開始
    wRen_S = Format(Val(wRen_S), "00")      '連 開始
    wRetu_E = Format(Val(wRetu_E), "00")    '列 終了
    wRen_E = Format(Val(wRen_E), "00")      '連 終了


    '棚マスタより指定範囲の有効データ検索
    Call UniCode_Conv(K0_TANA.Soko_No, wSoko)
    Call UniCode_Conv(K0_TANA.Retu, wRetu_S)
    Call UniCode_Conv(K0_TANA.Ren, wRen_S)
    Call UniCode_Conv(K0_TANA.Dan, "")
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(TANAREC.Soko_No, vbUnicode) <> wSoko Or _
                   StrConv(TANAREC.Retu, vbUnicode) > wRetu_E Or _
                   StrConv(TANAREC.Ren, vbUnicode) > wRen_E Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                Exit Function
        End Select

        '使用可能棚のみ処理対象
        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_OK Then

            '棚別個装箱マスタデータ書込み
            Call UniCode_Conv(TPACKINGREC.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.Retu, StrConv(TANAREC.Retu, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.Ren, StrConv(TANAREC.Ren, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.PACKING_NO, pPACKNo)
            Call UniCode_Conv(TPACKINGREC.RANK, pRANK)
            Call UniCode_Conv(TPACKINGREC.FILLER, "")
            sts = BTRV(BtOpInsert, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, com, "棚別個装箱マスタ")
                Exit Function
            End If

            '棚別個装箱マスタデータ生成件数表示
            wPUT_CNT = wPUT_CNT + 1
            L_CNT(plaData_Create).Caption = Format(wPUT_CNT, "#,0")
            DoEvents

        End If

        Call UniCode_Conv(K0_TANA.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_TANA.Retu, StrConv(TANAREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_TANA.Ren, StrConv(TANAREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_TANA.Dan, "99")
    Loop


    Data_Put_Proc = False


End Function
