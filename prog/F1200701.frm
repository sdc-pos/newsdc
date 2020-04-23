VERSION 5.00
Begin VB.Form F1200701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "月平均出荷数集計処理"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   6720
      MaxLength       =   4
      TabIndex        =   20
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   18
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1560
      Width           =   855
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
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
      Index           =   8
      Left            =   7800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      TabIndex        =   8
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実　行"
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label labMesg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "集計処理実行中です。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   23
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   21
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月～"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   19
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "集計年月"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
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
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   3720
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1200701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0               '国内外

Dim AVE_SYUKA_DATA   As String              '月平均出荷数データ

Dim MM_AVE          As Integer

Private Function OUTPUT_Proc() As Integer

Dim com         As Integer
Dim sts         As Integer
Dim Ret         As Integer

Dim FileNo      As Integer
Dim fileName    As String
    
Dim Data_cnt    As Integer
    
    
    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock             '画面項目ロック

    FileNo = FreeFile
    fileName = AVE_SYUKA_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo
    
    Write #FileNo, "[" & LabJIGYO.Caption & "]", "[" & Left(Combo(pcmbNAIGAI).Text, 2) & "]"
    
    Write #FileNo, "標準棚番", "品番（外）", "月平均出荷数"
    
    Data_cnt = 0
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                                            '範囲オーバー
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
                                        
                                        
        If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
        Else
                                        '標準棚番
            Write #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) + "-" & StrConv(ITEMREC.ST_DAN, vbUnicode),
                                        '品番（外）
            Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
                                        '月平均出荷数
            Write #FileNo, Format(CDbl(StrConv(ITEMREC.AVE_SYUKA, vbUnicode)), "#0.0")
        End If
    
    
        Data_cnt = Data_cnt + 1
    
        com = BtOpGetNext
    
    Loop

'    Write #FileNo, "品目件数＝" & Format(Data_cnt, "#0") & "件"


    Close #FileNo
    Call Input_UnLock             '画面項目ロック解除
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    OUTPUT_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1200701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200701)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200701)


    F1200701.MousePointer = vbDefault

End Sub
Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 0
            
            If Not IsNumeric(Text(1).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            Else
                Text(1).Text = Format(CInt(Text(1).Text), "00")
            End If
            
            If Not IsNumeric(Text(3).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            Else
                Text(3).Text = Format(CInt(Text(3).Text), "00")
            End If
            
            
            If (Text(0).Text & Text(1).Text) > (Text(2).Text & Text(3).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Exit Sub
            End If
            
                        
            
            
            Beep
            ans = MsgBox("「月平均出荷数」データ集計を実行しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If SUM_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
        
        Case 7                              'データ
            
            Beep
            ans = MsgBox("「月平均出荷数」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
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
                                '月平均出荷数ファイル名取り込み
    If GetIni("FILE", "AVE_SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "月平均出荷数データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    AVE_SYUKA_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1200701.Caption = "月平均出荷数集計処理（" + RTrim(JGYOBU_T(i).NAME) + ")"
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
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '画面初期設定
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbNAIGAI).SetFocus
    
    
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
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1200701 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).Code = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1200701.Caption = "月平均出荷数集計処理（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).Code
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
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(Index).SetFocus
                Exit Sub
            End If
     
            If Index = 1 Or Index = 3 Then
                Text(Index).Text = Format(CInt(Text(Index).Text), "00")
            End If
     
            
     
            For i = Index + 1 To 3
                If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                   Text(i).SetFocus
                   Exit For
                End If
            Next i
    End Select

End Sub

Private Function SUM_Proc() As Integer

Dim sts         As Integer
Dim Item_com    As Integer
Dim IDO_com     As Integer
Dim Item_Cnt    As Integer


Dim Sum_Syuka   As Long

    SUM_Proc = True

Label2(0).Caption = Format(Now, "HH:MM:SS")
    
    Call Input_Lock
    
    MM_AVE = DateDiff("m", (Text(0).Text & "/" & Text(1).Text & "/" & "01"), (Text(2).Text & "/" & Text(3).Text & "/" & "01"))
    MM_AVE = MM_AVE + 1

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    Item_com = BtOpGetGreaterEqual
    
    labMesg(0).Visible = True
    
    Item_Cnt = 0
    
    Do
        DoEvents
        
        Do
        
            sts = BTRV(Item_com + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                                            '範囲オーバー
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents
                Case Else
                    Call File_Error(sts, Item_com, "品目マスタ")
                    Exit Function
            End Select
        Loop
        If sts = BtErrEOF Then
            Exit Do
        End If

        Item_Cnt = Item_Cnt + 1
        Label1.Caption = Format(Item_Cnt, "#0")

        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_IDO.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.JITU_DT, Text(0).Text & Text(1).Text & "01")
        Call UniCode_Conv(K1_IDO.JITU_TM, "")
    
        IDO_com = BtOpGetGreaterEqual
    
        Sum_Syuka = 0
    
        Do
            DoEvents
            sts = BTRV(IDO_com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)

            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(IDOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Or _
                        Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                            '範囲オーバー
                        Exit Do
                    End If
                                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Text(2).Text & Text(3).Text & "31" Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents
                Case Else
                    Call File_Error(sts, Item_com, "品目マスタ")
                    Exit Function
            End Select
        
            If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Then
                If Right(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) <> CYU_KBN_BOU Then
        
                    Sum_Syuka = Sum_Syuka + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))

                
                End If
        
            End If
        
            IDO_com = BtOpGetNext
        
        Loop
        
        
        If MM_AVE = 0 Then
            Call UniCode_Conv(ITEMREC.AVE_SYUKA, "000000.0")
        Else
            Call UniCode_Conv(ITEMREC.AVE_SYUKA, Format(CDbl(Sum_Syuka / MM_AVE), "000000.0"))
        End If
        Do
        
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpUpdate, "品目マスタ")
                    Exit Function
            End Select
        
        Loop
    
        Item_com = BtOpGetNext
    
    Loop
Label2(1).Caption = Format(Now, "HH:MM:SS")

    MsgBox "正常終了しました！！"

    labMesg(0).Visible = False

    Call Input_UnLock

    SUM_Proc = False

End Function
