VERSION 5.00
Begin VB.Form F1060401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "棚番バーコード印刷"
   ClientHeight    =   6315
   ClientLeft      =   2025
   ClientTop       =   2655
   ClientWidth     =   11295
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
   ScaleHeight     =   6315
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷中止"
      Height          =   375
      Left            =   9000
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印  刷"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "確  定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   21
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Bar"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   28.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "F1060401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRT_CAN As Boolean                  '印刷途中キャンセル要求
Dim NormalFont As New StdFont           '印刷フォント
Dim Code39Font As New StdFont           '印刷フォント

Dim S_Tana As String
Dim E_Tana As String

Dim Text_Max As Integer                 '画面項目別最大ｲﾝﾃﾞｯｸｽ
Dim Command_Max As Integer
Private Sub Clear_Field()
Dim i As Integer

    For i = 0 To Text_Max
        Text(i).Text = ""
    Next i
End Sub


Private Function Print_Proc() As Integer

Dim sts As Integer
Dim flg As Boolean
Dim com As Integer

    Print_Proc = False


    PRT_CAN = False
    flg = False
    Call UniCode_Conv(K0_TANA.Soko_No, Text(0).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(1).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(2).Text)
    Call UniCode_Conv(K0_TANA.Dan, Text(3).Text)
    
    
        
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                Beep
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Unload Me
        End Select
                                            '明細印刷
                                            '印刷フォント設定
'        Printer.Print
'        Printer.Print
'        Printer.Print
        Printer.Print
        Set Printer.Font = Code39Font
        Printer.Print Tab(4);
        Printer.Print "*/" + StrConv(TANAREC.Soko_No, vbUnicode) + StrConv(TANAREC.Retu, vbUnicode) + StrConv(TANAREC.Ren, vbUnicode) + StrConv(TANAREC.Dan, vbUnicode) + "*"
        Printer.Print
        
        Set Printer.Font = NormalFont
        Printer.Print Tab(6);
        Printer.Print "*/" + StrConv(TANAREC.Soko_No, vbUnicode) + StrConv(TANAREC.Retu, vbUnicode) + StrConv(TANAREC.Ren, vbUnicode) + StrConv(TANAREC.Dan, vbUnicode) + "*"
        flg = True
        
    
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.Print
        Printer.Print
        Printer.Print


    PRT_CAN = False

End Function

Private Sub Command_Click(Index As Integer)

Dim yn As Integer
Dim RetBuf As String
Dim i As Integer
    
    Select Case Index
        Case 0                              '印刷
                                            'エラーチェック
            If Len(Text(0).Text) = 0 Then
                S_Tana = Space(8)
            Else
                If Len(Text(0).Text) <> 0 Then
                    S_Tana = Text(0).Text
                    For i = 1 To 3
                        If Not IsNumeric(Text(i).Text) Then
                            Beep
                            MsgBox "入力した項目はエラーです。", vbOKOnly + vbExclamation
                            Text(0).SelStart = 0
                            Text(0).SelLength = Len(Text(0).Text)
                            Text(0).SetFocus
                            Exit Sub
                        Else
                            S_Tana = S_Tana & Format(CInt(Text(i).Text), "00")
                            Text(i).Text = Format(CInt(Text(i).Text), "00")
                        End If
                    Next i
                Else
                    S_Tana = Space(8)
                End If

            End If

            Beep
            yn = MsgBox("確定しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(0).SelStart = 0
            Text(0).SelLength = Len(RTrim(Text(0).Text))
            Text(0).SetFocus
        Case 8                              '印刷
            Unload Me
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
    PRT_CAN = True
End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    Text_Max = 3                '画面項目別最大ｲﾝﾃﾞｯｸｽ
    Command_Max = 11

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    
                                '棚マスタＯＰＥＮ
    If TANA_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '印刷フォント設定
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '印刷フォント設定
    With NormalFont
        .NAME = F1060401.FontName
        .Size = 14
    End With
'    Set Printer.Font = NormalFont
                                
                                '画面初期設定
    Call Clear_Field
    
    Text(0).SelStart = 0
    Text(0).SelLength = Len(RTrim(Text(0).Text))
    Text(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
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
        Call File_Error(sts, BtOpReset, "棚マスタ")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub


Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If Index <> 0 Then
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Text(Index).SelStart = 0
                    Text(Index).SelLength = Len(Text(Index).Text)
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
            For i = Index + 1 To Text_Max
                If Text(i).Enabled Then
                    Text(i).SelStart = 0
                    Text(i).SelLength = Len(RTrim(Text(i).Text))
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled Then
                    Text(i).SelStart = 0
                    Text(i).SelLength = Len(RTrim(Text(i).Text))
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub

