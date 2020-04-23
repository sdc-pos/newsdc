VERSION 5.00
Begin VB.Form F1020701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "前借りリスト"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2265
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "印刷中断"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
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
      Index           =   11
      Left            =   10320
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
      Index           =   10
      Left            =   9480
      TabIndex        =   10
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
      TabIndex        =   9
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
      Index           =   8
      Left            =   7800
      TabIndex        =   8
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
      TabIndex        =   7
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
      Index           =   5
      Left            =   4800
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
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
      TabIndex        =   14
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷中です"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "F1020701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRT_CAN As Boolean                  '印刷途中キャンセル用
Dim NormalFont As New StdFont           '印刷フォント

Const LMAX% = 46                        '頁内最大行数
Const MGN_L% = 4                        '明細印刷開始桁位置（１から）
Const MGN_U% = 1                        '上余白（行数：１から）
Dim Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private Sub Print_Proc()

Dim Lcnt            As Integer
Dim sts             As Integer
Dim i               As Integer
Dim SV_JGYOBU       As String * 1
Dim Prt_JGYOBU      As String * 1
Dim RetBuf          As String
Dim com             As Integer
    
    
'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock           '画面項目ロック
    Label1.Visible = True
    DoEvents
    Command1.Visible = True
    DoEvents
    Command1.Enabled = True


    PRT_CAN = False

    
    Lcnt = 99
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time
    SV_JGYOBU = ""
    
    
    com = BtOpGetFirst
    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Exit Sub
        End If
        
        sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入荷チェックファイル（前借り）")
                Exit Sub
        End Select
        
        If SV_JGYOBU = " " Then
            SV_JGYOBU = StrConv(J_NYUREC.JGYOBU, vbUnicode)
        End If
        
        If SV_JGYOBU <> StrConv(J_NYUREC.JGYOBU, vbUnicode) Then
            SV_JGYOBU = StrConv(J_NYUREC.JGYOBU, vbUnicode)
            Lcnt = LMAX + 1
        End If
                                        'ヘッダー
        If (Lcnt > LMAX) Then
            Call Print_Head(Lcnt)
            Prt_JGYOBU = ""
        End If
    
        Printer.Print
        Printer.Print Tab(MGN_L + 3);
        If Prt_JGYOBU <> StrConv(J_NYUREC.JGYOBU, vbUnicode) Then
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(J_NYUREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    Printer.Print RTrim(JGYOBU_T(i).NAME);
                End If
            Next i
        End If
        
        Prt_JGYOBU = StrConv(J_NYUREC.JGYOBU, vbUnicode)
        Printer.Print Tab(MGN_L + 20);
        If StrConv(J_NYUREC.NAIGAI, vbUnicode) = NAIGAI_GAI Then
            Printer.Print NAIGAI2;
        Else
            Printer.Print "　　";
        End If
        Printer.Print "  " + StrConv(J_NYUREC.HIN_GAI, vbUnicode);
        
        Printer.Print Tab(MGN_L + 46);
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(J_NYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(J_NYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(J_NYUREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Printer.Print StrConv(ITEMREC.HIN_NAME, vbUnicode);
                Printer.Print Tab(MGN_L + 86);
                Printer.Print Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 13);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub
        End Select
        
        Printer.Print Tab(MGN_L + 116);
        RetBuf = Format(CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)), "#,##0")
        If Len(Trim(RetBuf)) < 10 Then
        
            RetBuf = Space(10 - Len(Trim(RetBuf))) & RetBuf
        End If
        
        
        Printer.Print RetBuf

        Lcnt = Lcnt + 2
        
        
        com = BtOpGetNext
    
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If
    




End Sub

Private Sub Print_Head(Lcnt As Integer)
                                        
Dim i       As Integer
Dim RetBuf  As String
Dim sts     As Integer

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(26);
    Printer.Print "＊＊＊  前借りリスト  ＊＊＊";
    Printer.Print Tab(85);
    Printer.Print "本日" & Pdate; "分" & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        
                                        '明細印刷
    Printer.Print Tab(MGN_L + 3);
    Printer.Print "事業部";
    Printer.Print Tab(MGN_L + 26);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 46);
    Printer.Print "品  名  ";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 118);
    Printer.Print "前借り数";
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub


Private Sub Command1_Click()
    PRT_CAN = True
End Sub

Private Sub Form_Activate()
Dim Ans As Integer
    
    Beep
    Ans = MsgBox("「前借りリスト」印刷しますか？", vbYesNo)
    If Ans = vbYes Then
        Call Print_Proc
    End If
    
    Unload Me

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

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)


                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '品目マスタOPEN
    If ITEM_Open(0) Then
        Unload Me
    End If
                                '入荷チェックファイルOPEN
    If J_NYU_Open(0) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020701.FontName
        .Size = F1020701.FontSize
    End With
    Set Printer.Font = NormalFont
    
    
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
                                            '入荷チェックファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷チェックファイル")
        End If
    End If
    
    sts = BTRV(BtOpReset, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020701 = Nothing

    End
End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020701)


    F1020701.MousePointer = vbDefault

End Sub

