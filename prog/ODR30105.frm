VERSION 5.00
Begin VB.Form ODR30105 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "希望納期　一括登録"
   ClientHeight    =   2115
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5445
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   225
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   75
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更　新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Lab_Dsp 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   675
      Width           =   3255
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "希望納期"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   675
      TabIndex        =   4
      Top             =   225
      Width           =   960
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
      End
   End
End
Attribute VB_Name = "ODR30105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'コンボ用添字
'Private Const pcmbHBUN = 0

'希望納期の最大先行月
Private Const Max_Day = 12


'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 0

Private Const ptxKIBOU_DT% = 0

'ラベル用添字
Private Const plabMSG% = 0

'コマンドボタン用添字
Private Const FuncCOR% = 0       '更新
Private Const FuncEND% = 1       '終了

'ListBox添字
'Private Const plst_DISP% = 0     '表示用データ　Sort順＆Key



Dim Init_F      As Integer


Private Function ERR_CHK(Index As Integer)
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String
Dim W_Date      As String

Dim W_Day       As Long

    ERR_CHK = True
    
                        '入力文字数チェック
    'If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
    '    MsgBox "入力した項目は（桁あふれエラー）です。", vbExclamation
    '    Exit Function
    'End If
    
    Select Case Index
        Case ptxKIBOU_DT%
            Lab_Dsp(plabMSG%) = ""
            W_STR = Trim(Text1(Index))
            
            If W_STR = "" Then
                MsgBox "希望納期　未設定！", vbExclamation
                Exit Function
            
            Else
            
                If Not IsDate(W_STR) Then
                    MsgBox "日付エラー！", vbExclamation
                    Exit Function
                
                End If
                
                W_STR = Format(Trim(Text1(Index)), "yyyy/mm/dd")
                Text1(Index) = W_STR
                DoEvents
                W_Date = Format(Date, "yyyy/mm/dd")
                
                If W_STR < W_Date Then
                    MsgBox "希望納期 ＜ 本日エラー！", vbExclamation
                    Exit Function
                
                End If
                
                W_Day = DateDiff("m", W_Date, W_STR)
                
                If W_Day > Max_Day Then
                
                    MsgBox Max_Day & "ケ月以上先エラー！", vbExclamation
                    
                    Exit Function
                End If
            
        End If
            
            
    End Select
    
    
    ERR_CHK = False
End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30105.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30105)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30105)


    ODR30105.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer
Dim X_i     As Integer
Dim W_After     As String

    Select Case Index
    
        Case FuncCOR%
            
            If ERR_CHK(ptxKIBOU_DT) Then
                Text1(ptxKIBOU_DT).SetFocus
                Call Text1_GotFocus(ptxKIBOU_DT)
                Exit Sub
            End If
            
            
            yn = MsgBox("更新しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            'yn = vbYes
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '更新処理
            KIBOU_DT = Text1(ptxKIBOU_DT)
            
            Init_F = 0
            ODR30105_Return = False                '確認画面 更新＆終了
            Me.Visible = False
            Exit Sub
            
        Case FuncEND%
            If ODR30105_Return = True Then
                'yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
                yn = vbYes
                
                If yn = vbNo Then
                
                    Exit Sub
                End If
            End If
            
            Init_F = 0
            ODR30105_Return = True                '確認画面ｷｬﾝｾﾙ終了
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR30105.Top = ODR30101.Top + (ODR30101.Height - ODR30105.Height) / 2
    
    
    
    ODR30105.Left = ODR30101.Left + (ODR30101.Width - ODR30105.Width) / 2
    
    
    
    
    Text1(ptxKIBOU_DT).SetFocus
    Call Text1_GotFocus(ptxKIBOU_DT)
    
    ODR30105_Return = True
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()
Dim cc As tagINITCOMMONCONTROLSEX
'Dim PanePos(2) As Long

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String




'コモンコントロールを初期化する
cc.dwSize = Len(cc)
cc.dwICC = ICC_BAR_CLASSES

'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "希望納期登録", Me.hwnd, 0)
'ペイン複数作る
'最後の要素を-1にすると
'親ウィンドウの全体の幅の残りの幅を
'自動的に割り当てる
'PanePos(0) = 200
'PanePos(1) = 300
'PanePos(2) = -1
'Call SendMessageAny(hStatusWnd, SB_SETPARTS, 3, PanePos(0))
Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


'画面初期処理
    'Show
    
    'Text1(ptxTANTO_CD).SetFocus
    'Max_Row = 25000
    
    Init_F = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode <> 0 Then Exit Sub
    'If UnloadMode = 1 Then Exit Sub
    
    yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Me.Visible = False
    
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
    
        Case 0      '更新
            Call Command1_Click(FuncCOR)
        
        
        Case 1       '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub


Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index))
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index))
    End If
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Text1(Index).Locked = True Then      'ロック中項目なら処理しない
        Call Tab_Ctrl(Shift)    '移動
        Exit Sub
    End If
                        '入力文字数チェック
    If ERR_CHK(Index) Then
        Call Text1_GotFocus(Index)
        Text1(Index).SetFocus
        Exit Sub
    End If
    
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub

