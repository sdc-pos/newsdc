VERSION 5.00
Begin VB.Form F1060701 
   BackColor       =   &H00C0C0C0&
   Caption         =   "欠品防止支援処理"
   ClientHeight    =   4710
   ClientLeft      =   2025
   ClientTop       =   2265
   ClientWidth     =   7320
   ControlBox      =   0   'False
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "実行中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "欠品防止支援処理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3840
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
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1060701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private KeppinArarm_DATA    As String       '欠品アラームデータフルパス
Private Kakeritu            As Integer      '欠品防止掛け率

Private Function OUTPUT_Proc() As Integer
'----------------------------------------------------------------------------
'                  ＣＳＶデータ出力処理
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim FileNo          As Integer
Dim fileName        As String


Dim AVE_SYUKA       As Long

Dim Alarm_Flg       As Boolean

Dim c               As String * 128
Dim Soko_No         As String * 2


    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    Label1(0).Visible = True
    Label1(1).Visible = True
    

    '-------------------------------------------    欠品防止ログから増加が有った品目を削除する
    com = BtOpGetFirst
    Do
        DoEvents
        '在庫集計データ読み込み
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫集計データ")
                Exit Function
        End Select
            
        If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) > CLng(StrConv(SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) Then
            '在庫増加が有ったら月平均出荷数読み込み
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
    
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    AVE_SYUKA = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                Case BtErrKeyNotFound
                    AVE_SYUKA = 0
                Case Else
                    Call File_Error(sts, com, "在庫集計データ")
                    Exit Function
            End Select
                '前日在庫数　＞　月平均出荷数　*　ｎ％
            If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) > (AVE_SYUKA * (Kakeritu / 100)) Then
                '欠品防止ログから消去する
                Call UniCode_Conv(K0_KEPPINLOG.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
               
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            'この処理では本来ありえない！！
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "欠品防止支援ログ")
                            Exit Function
                    End Select
                Loop
        
                If sts = BtNoErr Then
        
                    Do
                        sts = BTRV(BtOpDelete, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                'この処理では本来ありえない！！
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                           Case Else
                                Call File_Error(sts, BtOpDelete, "欠品防止支援ログ")
                                Exit Function
                        End Select
                    Loop
        
                End If
            End If
        End If
    
        com = BtOpGetNext
    
    Loop
    '-------------------------------------------
    
    FileNo = FreeFile
    fileName = KeppinArarm_DATA
    
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo


    Write #FileNo, "欠品防止支援リスト(" & Format(Now, "YYYY/MM/DD") & "作成）"
    Write #FileNo, "事業部", "国内外", "品番（外部）", "理論在庫", "月平均出荷数", "標準棚番"
    
    
    '-------------------------------------------    前々日から在庫が減った分の欠品をチェックする
    Alarm_Flg = False
    com = BtOpGetFirst
    Do
        DoEvents
        '在庫集計データ読み込み
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫集計データ")
                Exit Function
        End Select
    
        If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) < CLng(StrConv(SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) Then
        
            '在庫減少が有ったら月平均出荷数読み込み
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
    
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    AVE_SYUKA = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                Case BtErrKeyNotFound
                    AVE_SYUKA = 0
                Case Else
                    Call File_Error(sts, com, "在庫集計データ")
                    Exit Function
            End Select
        
            If CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) < (AVE_SYUKA * (Kakeritu / 100)) Then
                '在庫数が少なくなったら欠品防止ログをチェックする
                Call UniCode_Conv(K0_KEPPINLOG.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_KEPPINLOG.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
               
                sts = BTRV(BtOpGetEqual, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "欠品防止支援ログ")
                        Exit Function
                End Select
                            
            
            
                If sts = BtErrKeyNotFound Then
                    '欠品防止ログ未登録なら
                    Alarm_Flg = True
                
                                                '事業部
                    Write #FileNo, StrConv(SUMZREC.JGYOBU, vbUnicode),
                                                '国内外
                    Write #FileNo, StrConv(SUMZREC.NAIGAI, vbUnicode),
                                                '品目（外部）
                    Write #FileNo, StrConv(SUMZREC.HIN_GAI, vbUnicode),
                                                '前日理論在庫
                    Write #FileNo, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))),
                                                '前日理論在庫
                    Write #FileNo, Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))),
                                                '標準棚番
                    If GetIni("SOKO_NO", StrConv(SUMZREC.ST_SOKO, vbUnicode), "SYS", c) Then
                        Soko_No = StrConv(SUMZREC.ST_SOKO, vbUnicode)
                    Else
                        Soko_No = Trim(c)
                    End If
                    
                    
                    
                    Write #FileNo, Soko_No & "-" _
                                     & StrConv(SUMZREC.ST_RETU, vbUnicode) & "-" _
                                     & StrConv(SUMZREC.ST_REN, vbUnicode) & "-" _
                                     & StrConv(SUMZREC.ST_DAN, vbUnicode)
                
                    '欠品ログ出力
                                    
                    Call UniCode_Conv(KEPPINLOGREC.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(KEPPINLOGREC.CREATE_DT, Format(Now, "YYYYMMDD"))
                    Do
                        sts = BTRV(BtOpInsert, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                'この処理では本来ありえない！！
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                           Case Else
                                Call File_Error(sts, BtOpInsert, "欠品防止支援ログ")
                                Exit Function
                        End Select
                    Loop

                
                
                
                End If
            
            
            End If
        End If
    
        com = BtOpGetNext
    Loop




    Close #FileNo
    
    Call Input_UnLock         '画面項目ロック解除
    
    If Alarm_Flg Then
        Beep
        MsgBox "欠品防止の対象品目が有りました。「" & fileName & "」が出力されました。"
    Else
        Beep
        MsgBox "欠品防止の対象品目は有りませんでした。"
    End If
    
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

    F1060701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060701)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060701)


    F1060701.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

Dim ans     As Integer


    ans = MsgBox("「欠品防止支援処理」実行しますか？", vbYesNo, "確認処理")
    
    If ans = vbYes Then
        If OUTPUT_Proc() Then
            Unload Me
        End If
    End If

    Unload Me

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
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

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
    LOG_F = Trim(c)
                                '欠品防止支援データファイル名取り込み
    If GetIni("FILE", "KeppinArarm_DATA", "SYS", c) Then
        Beep
        MsgBox "欠品防止支援データ作成ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KeppinArarm_DATA = Trim(c)
                                '欠品防止掛け率
    If GetIni(App.EXEName, "KAKERITU", "SYS", c) Then
        Kakeritu = 100
    Else
        If IsNumeric(Trim(c)) Then
            Kakeritu = CInt(Trim(c))
        Else
            Kakeritu = 100
        End If
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '欠品防止支援ログＯＰＥＮ
    If KEPPINLOG_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
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
                                            '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
                                            '月平均出荷数ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷数")
        End If
    End If
                                            '欠品防止支援ログＣＬＯＳＥ
    sts = BTRV(BtOpClose, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "欠品防止支援ログ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060701 = Nothing

    End
End Sub
