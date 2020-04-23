VERSION 5.00
Begin VB.Form F3010201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GLICS在庫対応　棚別在庫一覧表印刷"
   ClientHeight    =   6840
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   11430
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   6720
      MaxLength       =   13
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   3840
      MaxLength       =   13
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   0
      Top             =   360
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
      TabIndex        =   21
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
      Index           =   10
      Left            =   9480
      TabIndex        =   20
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
      Index           =   9
      Left            =   8640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印 刷"
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
      TabIndex        =   18
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
      Index           =   7
      Left            =   6480
      TabIndex        =   17
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
      Index           =   6
      Left            =   5640
      TabIndex        =   16
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
      Index           =   5
      Left            =   4800
      TabIndex        =   15
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
      Index           =   4
      Left            =   3960
      TabIndex        =   14
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
      Index           =   3
      Left            =   2640
      TabIndex        =   13
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
      Index           =   2
      Left            =   1800
      TabIndex        =   12
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
      TabIndex        =   11
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   240
      Index           =   10
      Left            =   2280
      TabIndex        =   32
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   9
      Left            =   5880
      TabIndex        =   31
      Top             =   2880
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "　　　段"
      Height          =   240
      Index           =   8
      Left            =   2280
      TabIndex        =   30
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   7
      Left            =   5880
      TabIndex        =   29
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "　　　連"
      Height          =   240
      Index           =   6
      Left            =   2280
      TabIndex        =   28
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   5
      Left            =   5880
      TabIndex        =   27
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番　列"
      Height          =   240
      Index           =   4
      Left            =   2280
      TabIndex        =   26
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   3
      Left            =   5880
      TabIndex        =   25
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷中です"
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   1
      Left            =   5880
      TabIndex        =   23
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫№"
      Height          =   240
      Index           =   0
      Left            =   2280
      TabIndex        =   22
      Top             =   480
      Width           =   720
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F3010201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_SOKO_NO% = 0          '開始　倉庫№
Private Const ptxE_SOKO_NO% = 1             '終了　倉庫№
Private Const ptxS_RETU% = 2                '開始　棚番　列
Private Const ptxE_RETU% = 3                '終了　棚番　列
Private Const ptxS_REN% = 4                 '開始　棚番　連
Private Const ptxE_REN% = 5                 '終了　棚番　連
Private Const ptxS_DAN% = 6                 '開始　棚番　段
Private Const ptxE_DAN% = 7                 '終了　棚番　段
Private Const ptxS_HIN_GAI% = 8             '開始　品番（外部）
Private Const ptxE_HIN_GAI% = 9             '開始　品番（外部）



Private Const Text_Max% = 9                 '画面項目別最大ｲﾝﾃﾞｯｸｽ


Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim LCNT            As Integer
'Dim PRI_Location    As String

Dim Pdate   As String                       '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime   As String                       '印刷開始時刻（ﾍｯﾀﾞｰ用）

Dim NormalFont As New StdFont               '印刷フォント




Private Function Print_Proc() As Integer

Dim com             As Integer
Dim sts             As Integer



Dim SAVE_HIN_GAI    As String
Dim SAVE_Location   As String

    
    
    Print_Proc = True
'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock           '画面項目ロック
    Label1.Visible = True

'印刷開始
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time

    LCNT = 99
    
    SAVE_HIN_GAI = ""
    SAVE_Location = ""
    
    

    Call UniCode_Conv(K0_ZAIKO.Soko_No, Text(ptxS_SOKO_NO).Text)
    Call UniCode_Conv(K0_ZAIKO.Retu, Text(ptxS_RETU).Text)
    Call UniCode_Conv(K0_ZAIKO.Ren, Text(ptxS_REN).Text)
    Call UniCode_Conv(K0_ZAIKO.Dan, Text(ptxS_DAN).Text)
    Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
     
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                    StrConv(ZAIKOREC.Retu, vbUnicode) & _
                    StrConv(ZAIKOREC.Ren, vbUnicode) & _
                    StrConv(ZAIKOREC.Dan, vbUnicode)) > _
                    (Text(ptxE_SOKO_NO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                
                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) < Text(ptxS_HIN_GAI).Text Or _
            StrConv(ZAIKOREC.HIN_GAI, vbUnicode) > Text(ptxE_HIN_GAI).Text Then
        Else
            If Trim(SAVE_Location) = "" Then
                SAVE_Location = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                StrConv(ZAIKOREC.Dan, vbUnicode)
                SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            End If
        
            If SAVE_Location <> StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                StrConv(ZAIKOREC.Dan, vbUnicode) Then
                If Print_Sub_Proc(SAVE_Location, SAVE_HIN_GAI) Then
                    Exit Function
                End If
                SAVE_Location = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                StrConv(ZAIKOREC.Dan, vbUnicode)
                SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            
            End If
        
            If SAVE_HIN_GAI <> StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
                If Print_Sub_Proc(SAVE_Location, SAVE_HIN_GAI) Then
                    Exit Function
                End If
                SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            End If
        
        End If
    
    
    
    Loop
    
    If Trim(SAVE_Location) <> "" Then
        If Print_Sub_Proc(SAVE_Location, SAVE_HIN_GAI) Then
            Exit Function
        End If
    End If

    If LCNT <> 99 Then
        Printer.EndDoc
    End If

    Call Input_UnLock               '画面項目ロック解除

    Print_Proc = False

End Function
Private Sub Print_Head()
'ヘッダ印刷
Dim i As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(36);
    Printer.Print "＊＊＊  棚別在庫一覧表  ＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print

                                        'ヘッダー（３）
    Printer.Print Tab(MGN_L + 2);
    Printer.Print "棚　番";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "品番(外部)";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "品　　名";
    Printer.Print Tab(MGN_L + 56);
    Printer.Print "在庫数";
    Printer.Print Tab(MGN_L + 66);
    Printer.Print "(完)　(未)";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "別置(棚番)";
    Printer.Print Tab(MGN_L + 97);
    Printer.Print "別置(在庫)";
    Printer.Print Tab(MGN_L + 108);
    Printer.Print "Glics(S2)";
    Printer.Print Tab(MGN_L + 118);
    Printer.Print "Glics(P2)"
    
    
    Printer.Print

    Printer.Print

    LCNT = 7 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F3010201.MousePointer = vbHourglass

    Call Ctrl_Lock(F3010201)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F3010201)


    F3010201.MousePointer = vbDefault

End Sub




Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「棚別在庫一覧表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
End Sub
Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select


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
    

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F3010201.Caption = "GLICS在庫対応　棚別在庫一覧表印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データ(ﾜｰｸ)ＯＰＥＮ
    If wZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F3010201.FontName
        .Size = F3010201.FontSize
    End With
    Set Printer.Font = NormalFont
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫データ(ﾜｰｸ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ(ﾜｰｸ)")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F3010201 = Nothing

    End
End Sub
Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F3010201.Caption = "GLICS在庫対応　棚別在庫一覧表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If



End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub



Private Function Err_Chk()
    
Dim i As Integer
    
    Err_Chk = True

'倉庫番号

    If Len(Text(ptxE_SOKO_NO).Text) = 0 Then
        Text(ptxE_SOKO_NO).Text = "zz"
    End If

    If Text(ptxS_SOKO_NO).Text > Text(ptxE_SOKO_NO).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_SOKO_NO).SetFocus
        Exit Function
    End If

'棚番
    For i = ptxS_RETU To ptxE_DAN
        Select Case i
            Case ptxS_RETU, ptxS_REN, ptxS_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "00"
                End If
            Case ptxE_RETU, ptxE_REN, ptxE_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "99"
                End If
        End Select
        If IsNumeric(Text(i).Text) Then
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i


    If Text(ptxS_RETU).Text & Text(ptxS_REN).Text & Text(ptxS_DAN).Text _
        > Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_RETU).SetFocus
        Exit Function
    End If
'品番(外部)
    If Len(Text(ptxE_HIN_GAI).Text) = 0 Then
        Text(ptxE_HIN_GAI).Text = String(Len(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)), "z")
    End If

    If Text(ptxS_HIN_GAI).Text > Text(ptxE_HIN_GAI).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_HIN_GAI).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function

Private Function Print_Sub_Proc(Location As String, Hinban As String) As Integer


Dim sts             As Integer
Dim com             As Integer


Dim Tana_Zaiko_Qty  As Long
Dim Betu_Zaiko_Qty  As Long


Dim SAVE_Location   As String

Dim Print_Cnt       As Integer

    Print_Sub_Proc = True


    '品目マスタの読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            '有ったらおかしい！！
            Print_Sub_Proc = False
            Exit Function
        
        Case Else
            Call File_Error(sts, com, "在庫ﾃﾞｰﾀ")
            Exit Function
    End Select


    '現在の指定棚の在庫を検索
    Call UniCode_Conv(K4_wZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_wZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K4_wZAIKO.HIN_GAI, Hinban)
    
    Call UniCode_Conv(K4_wZAIKO.Soko_No, Mid(Location, 1, 2))
    Call UniCode_Conv(K4_wZAIKO.Retu, Mid(Location, 3, 2))
    Call UniCode_Conv(K4_wZAIKO.Ren, Mid(Location, 5, 2))
    Call UniCode_Conv(K4_wZAIKO.Dan, Mid(Location, 7, 2))

    Tana_Zaiko_Qty = 0

    com = BtOpGetGreaterEqual

    Do
        DoEvents
    
        sts = BTRV(com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K4_wZAIKO, Len(K4_wZAIKO), 4)
        Select Case sts
            Case BtNoErr
                
                If StrConv(wZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(wZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
            
                If StrConv(wZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
            
                If StrConv(wZAIKOREC.Soko_No, vbUnicode) & _
                    StrConv(wZAIKOREC.Retu, vbUnicode) & _
                    StrConv(wZAIKOREC.Ren, vbUnicode) & _
                    StrConv(wZAIKOREC.Dan, vbUnicode) <> Location Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "w在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        Tana_Zaiko_Qty = Tana_Zaiko_Qty + CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode))
                
        com = BtOpGetNext
    
    
    Loop


    SAVE_Location = ""
    Betu_Zaiko_Qty = 0
    Print_Cnt = 0


    Call UniCode_Conv(K4_wZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_wZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K4_wZAIKO.HIN_GAI, Hinban)
    
    Call UniCode_Conv(K4_wZAIKO.Soko_No, "")
    Call UniCode_Conv(K4_wZAIKO.Retu, "")
    Call UniCode_Conv(K4_wZAIKO.Ren, "")
    Call UniCode_Conv(K4_wZAIKO.Dan, "")


    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K4_wZAIKO, Len(K4_wZAIKO), 4)
        Select Case sts
            Case BtNoErr
                
                If StrConv(wZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(wZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
            
                If StrConv(wZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If (StrConv(wZAIKOREC.Soko_No, vbUnicode) & _
            StrConv(wZAIKOREC.Retu, vbUnicode) & _
            StrConv(wZAIKOREC.Ren, vbUnicode) & _
            StrConv(wZAIKOREC.Dan, vbUnicode)) = Location Then
        Else
    
    
                If Trim(SAVE_Location) = "" Then
                    SAVE_Location = StrConv(wZAIKOREC.Soko_No, vbUnicode) & _
                                    StrConv(wZAIKOREC.Retu, vbUnicode) & _
                                    StrConv(wZAIKOREC.Ren, vbUnicode) & _
                                    StrConv(wZAIKOREC.Dan, vbUnicode)
                End If
            
            
                If SAVE_Location <> StrConv(wZAIKOREC.Soko_No, vbUnicode) & _
                                    StrConv(wZAIKOREC.Retu, vbUnicode) & _
                                    StrConv(wZAIKOREC.Ren, vbUnicode) & _
                                    StrConv(wZAIKOREC.Dan, vbUnicode) Then
                
                    Call Detail_Print_Proc(Hinban, Location, SAVE_Location, Tana_Zaiko_Qty, Betu_Zaiko_Qty, Print_Cnt)
                
                    SAVE_Location = StrConv(wZAIKOREC.Soko_No, vbUnicode) & _
                                    StrConv(wZAIKOREC.Retu, vbUnicode) & _
                                    StrConv(wZAIKOREC.Ren, vbUnicode) & _
                                    StrConv(wZAIKOREC.Dan, vbUnicode)
                
                
                    Betu_Zaiko_Qty = 0
                
                    Print_Cnt = Print_Cnt + 1
                
                
                End If
            
            
                Betu_Zaiko_Qty = Betu_Zaiko_Qty + CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode))
        
        End If
    
        com = BtOpGetNext
    
    Loop

'    If Trim(SAVE_Location) <> "" Then
        Call Detail_Print_Proc(Hinban, Location, SAVE_Location, Tana_Zaiko_Qty, Betu_Zaiko_Qty, Print_Cnt)
    
        Printer.Print
        LCNT = LCNT + 1
    
'    End If
    Print_Sub_Proc = False

End Function

Private Sub Detail_Print_Proc(Hinban As String, Location As String, Batu_Location As String, Zaiko_Qty As Long, Batu_Zaiko_Qty As Long, Print_Cnt As Integer)

    If LCNT > LMAX Then
        Call Print_Head
        Print_Cnt = 0
'        PRI_Location = ""
    End If

    If Print_Cnt = 0 Then
        
        Printer.Print Tab(MGN_L);
        
        
'        If PRI_Location <> Location Then
            If Location = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                            StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                Printer.Print "*";
            End If
        
            Printer.Print Tab(MGN_L + 2);
        
            Printer.Print Mid(Location, 1, 2) & "-" _
                            & Mid(Location, 3, 2) & "-" _
                            & Mid(Location, 5, 2) & "-" & _
                            Mid(Location, 7, 2);
'            PRI_Location = Location
        
'        End If
    
        Printer.Print Tab(MGN_L + 15);
    
        Printer.Print Left(Hinban, 13);
    
        Printer.Print Tab(MGN_L + 30);
        Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
            
        Printer.Print Tab(MGN_L + 56);
                            
        Printer.Print Space(6 - Len(Format(Zaiko_Qty, "#,##0"))) & Format(Zaiko_Qty, "#,##0");
            
            
    End If

    Printer.Print Tab(MGN_L + 80);


    If Trim(Batu_Location) = "" Then
    Else
        If Batu_Location = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
            Printer.Print "*";
        End If
        
        Printer.Print Tab(MGN_L + 82);
        
        Printer.Print Mid(Batu_Location, 1, 2) & "-" _
                        & Mid(Batu_Location, 3, 2) & "-" _
                        & Mid(Batu_Location, 5, 2) & "-" & _
                        Mid(Batu_Location, 7, 2);

        Printer.Print Tab(MGN_L + 100);
                            
        Printer.Print Space(6 - Len(Format(Batu_Zaiko_Qty, "#,##0"))) & Format(Batu_Zaiko_Qty, "#,##0");
    End If

    If Print_Cnt = 0 Then
    
        Printer.Print Tab(MGN_L + 111);
        Printer.Print Space(6 - Len(Format(CLng(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)), "#,##0"))) & Format(CLng(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)), "#,##0");
    
        Printer.Print Tab(MGN_L + 121);
        Printer.Print Space(6 - Len(Format(CLng(StrConv(ITEMREC.G_P2_ZAI_QTY, vbUnicode)), "#,##0"))) & Format(CLng(StrConv(ITEMREC.G_P2_ZAI_QTY, vbUnicode)), "#,##0");
    End If

    Printer.Print
    LCNT = LCNT + 1
End Sub
