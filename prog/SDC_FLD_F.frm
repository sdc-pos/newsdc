VERSION 5.00
Begin VB.Form SDC_FLD_F 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "データ出力先指定／確認"
   ClientHeight    =   3120
   ClientLeft      =   30
   ClientTop       =   3300
   ClientWidth     =   6975
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   ScaleHeight     =   3120
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   4788
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   1
      Top             =   840
      Width           =   4788
   End
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2040
      Width           =   468
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   2
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   2
      Top             =   1440
      Width           =   4788
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   11
      Left            =   6000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   10
      Left            =   5520
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   5040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ＯＫ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   4200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   3600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   3120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   492
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      Caption         =   "フォルダ名"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      Caption         =   "識別子"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      Caption         =   "ファイル名"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Lab_Fix 
      AutoSize        =   -1  'True
      Caption         =   "出力先ルート"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "SDC_FLD_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'コンボ用添字
'Const pcmb = ZERO

'テキスト用添字
Const ptxROOT% = 0                      '出力先ルート
Const ptxFOLDER% = 1                     'フォルダ名
Const ptxFILE% = 2                       'ファイル名
Const ptxXXX% = 3                        '識別子

'ラベル用添字
'Const plb = ZERO

'リスト用添字
'Const plst = ZERO

'ファンクション添字
Const fncOK% = 8                         'ＯＫ
Const fncCAN% = 11                       'ｷｬﾝｾﾙ

Dim Act_Flg As Integer                  'Activateｺﾝﾄﾛｰﾙﾌﾗｸﾞ

Private Function Data_Chk_Set() As Integer
'----------------------------------------------------------------------------
'                   入力データまとめチェック　＆　取込み
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i As Integer
Dim Wk As String
Dim Er_Idx As Integer
Dim yn As Integer

    Data_Chk_Set = True

    For i = ptxFOLDER To ptxXXX
        If Text1(i).MaxLength <> 0 And _
           In_Chr_Chk(Text1(i), Text1(i).MaxLength) Then
            Er_Idx = i
            GoTo Err_Return
        End If
    Next i

    'フォルダ有無確認
    If Len(Trim(Text1(ptxFOLDER))) <> 0 Then
        Er_Idx = ptxFOLDER
        Wk = Trim(Text1(ptxROOT)) & "\" & Trim(Text1(ptxFOLDER))
        If Dir(Wk, vbDirectory) = "" Then
            yn = MsgBox("フォルダが存在しません" & Chr(13) & Chr(10) & _
                        "作成しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbNo Then GoTo Err_Return
            MkDir Wk
        End If
    End If

    '入力値取込み
    SDC_FLD_Root = Trim(Text1(ptxROOT))      '出力先ルート
    SDC_FLD_Folder = Trim(Text1(ptxFOLDER))  'フォルダ名
    SDC_FLD_File = Trim(Text1(ptxFILE))      'ファイル名
    SDC_FLD_xxx = Trim(Text1(ptxXXX))        '識別子

    Data_Chk_Set = False
    Exit Function

Err_Return:
    Text1(Er_Idx).SetFocus

End Function

Private Function In_Chr_Chk(DATA As String, C_Len As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力文字項目エラーチェック
'----------------------------------------------------------------------------

    In_Chr_Chk = True

    If LenB(StrConv(DATA, vbFromUnicode)) > C_Len Then
        Beep
        MsgBox "入力した項目はエラーです。（桁あふれ）", vbExclamation
        Exit Function
    End If

    In_Chr_Chk = False

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
    SDC_FLD_F.MousePointer = vbHourglass
    Call Ctrl_Lock(SDC_FLD_F)
End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
    Call Ctrl_UnLock(SDC_FLD_F)
    SDC_FLD_F.MousePointer = vbDefault
End Sub

Private Sub Command1_Click(Index As Integer)
'----------------------------------------------------------------------------
'                   処理ファンクション　コントロール（最下段の１２個）
'----------------------------------------------------------------------------
Dim yn As Integer

    Select Case Index

        Case fncOK              '出力ＯＫ
            If Data_Chk_Set Then Exit Sub
            SDC_FLD_Return = False               '確認画面ＯＫ終了

        Case fncCAN             '処理ｷｬﾝｾﾙ
            SDC_FLD_Return = True                '確認画面ｷｬﾝｾﾙ終了

        Case Else
            Exit Sub
    End Select

    Act_Flg = False                 'Activateｺﾝﾄﾛｰﾙﾌﾗｸﾞ
    SDC_FLD_F.Visible = False
End Sub

Private Sub Form_Activate()
Dim sts As Integer
Dim i As Integer
Dim yn As Integer

    If Act_Flg = True Then Exit Sub

    Act_Flg = True                      'Activateｺﾝﾄﾛｰﾙﾌﾗｸﾞ

    '出力先情報　初期表示
    Text1(ptxROOT) = SDC_FLD_Root        '出力先ルート
    Text1(ptxFOLDER) = SDC_FLD_Folder    'フォルダ名
    Text1(ptxFILE) = SDC_FLD_File        'ファイル名
    Text1(ptxXXX) = SDC_FLD_xxx          '識別子

    DoEvents            '画面項目表示

    '出力先ルートフォルダ有無確認
    If Dir(SDC_FLD_Root, vbDirectory) = "" Then
        yn = MsgBox("出力先ルートが存在しません" & Chr(13) & Chr(10) & _
                    "作成しますか？", vbYesNo + vbQuestion, "確認入力")
        If yn = vbNo Then
            MsgBox "出力先ルートを正しく定義してから" & Chr(13) & Chr(10) & _
                   "再起動して下さい"
            Command1(fncCAN).Value = True
            Exit Sub
        End If
        MkDir SDC_FLD_Root
    End If



    Text1(ptxFOLDER).SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            If Command1(KeyCode - vbKeyF1).Enabled = False Then Exit Sub
            Command1(KeyCode - vbKeyF1).Value = True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Act_Flg = False                             'Activateｺﾝﾄﾛｰﾙﾌﾗｸﾞ
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).Locked <> True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then Exit Sub
    'Call Tab_Ctrl(Shift)    '移動
    If Index < 3 Then
        Text1(Index + 1).SetFocus
        Call Text1_GotFocus(Index + 1)
    Else
        Command1(8).SetFocus
    End If
    

End Sub

