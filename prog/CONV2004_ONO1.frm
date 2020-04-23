VERSION 5.00
Begin VB.Form CONV2004_ONO1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理（炊飯⇒アイロン）"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '画面の中央
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label OUT_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　出荷予定＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "在庫移動歴＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label IN_Cnt 
      Alignment       =   2  '中央揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "在庫データ＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データコンバート処理（炊飯⇒アイロン）"
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
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   9120
   End
End
Attribute VB_Name = "CONV2004_ONO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim IN_Count        As Long
Dim OUT_Count       As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  品目マスタのコンバート
    
    Call Log_Out(LOG_F, "品目マスタコンバート開始=" & Format(Now, "HH:MM:SS"))
    
    
    MsgLab(1) = "品目マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0

    IN_Cnt(0).Caption = Format(IN_Count, "#0")
    OUT_Cnt(0).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_OLD_ITEM2, Len(K0_OLD_ITEM2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧２）品目マスタ")
                Exit Function
        End Select


        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(0).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If




        If StrConv(OLD_ITEM2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_ITEM2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_ITEM2REC.JGYOBU, SENTAKU)
                OUT_Count = OUT_Count + 1
            End If
        End If
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "品目マスタ" & "[" & StrConv(OLD_ITEM2REC.JGYOBU, vbUnicode) & "-" & StrConv(OLD_ITEM2REC.NAIGAI, vbUnicode) & "-" & StrConv(OLD_ITEM2REC.HIN_GAI, vbUnicode))
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop

    OUT_Cnt(0).Caption = Format(OUT_Count, "#0")    '炊飯⇒アイロン更新件数


'---------------------------------------------  在庫データのコンバート
    Call Log_Out(LOG_F, "在庫データコンバート開始=" & Format(Now, "HH:MM:SS"))
    
    
    MsgLab(1) = "在庫データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0


    IN_Cnt(1).Caption = Format(IN_Count, "#0")
    OUT_Cnt(1).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_ZAIKO2_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_OLD_ZAIKO2, Len(K0_OLD_ZAIKO2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧２）在庫データ")
                Exit Function
        End Select

        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(1).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If


        If StrConv(OLD_ZAIKO2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_ZAIKO2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_ZAIKO2REC.JGYOBU, SENTAKU)

                OUT_Count = OUT_Count + 1

            End If
        End If

        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫データ")
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop

    OUT_Cnt(1).Caption = Format(OUT_Count, "#0")    '炊飯⇒アイロン更新件数

'---------------------------------------------  在庫移動歴のコンバート
    Call Log_Out(LOG_F, "在庫移動歴コンバート開始=" & Format(Now, "HH:MM:SS"))
    
    MsgLab(1) = "在庫移動歴コンバート処理中！！"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0

    IN_Cnt(2).Caption = Format(IN_Count, "#0")
    OUT_Cnt(2).Caption = Format(OUT_Count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), K0_OLD_IDO2, Len(K0_OLD_IDO2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧２）在庫移動歴")
                Exit Function
        End Select


        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(2).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If


        If StrConv(OLD_IDO2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_IDO2REC.HIN_GAI, vbUnicode)), "SETUP", c)
            If Not sts Then

                Call UniCode_Conv(OLD_IDO2REC.JGYOBU, SENTAKU)

                OUT_Count = OUT_Count + 1

            End If

        End If
        Do
            sts = BTRV(BtOpInsert, IDO_POS, OLD_IDO2REC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫移動歴")
                    Exit Function
            End Select
        Loop



        com = BtOpGetNext

    Loop
    
    OUT_Cnt(2).Caption = Format(OUT_Count, "#0")    '炊飯⇒アイロン更新件数

'---------------------------------------------  出荷予定のコンバート
    Call Log_Out(LOG_F, "出荷予定コンバート開始=" & Format(Now, "HH:MM:SS"))
    
    MsgLab(1) = "出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    IN_Count = 0
    OUT_Count = 0
    DISP_INTERVAL = 0
    
    IN_Cnt(3).Caption = Format(IN_Count, "#0")
    OUT_Cnt(3).Caption = Format(OUT_Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU2_POS, OLD_Y_SYU2REC, Len(OLD_Y_SYU2REC), K0_OLD_Y_SYU2, Len(K0_OLD_Y_SYU2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧２）出荷予定データ")
                Exit Function
        End Select
        
        
        IN_Count = IN_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            IN_Cnt(3).Caption = Format(IN_Count, "#0")
            DISP_INTERVAL = 0
        End If

        
        If StrConv(OLD_Y_SYU2REC.JGYOBU, vbUnicode) = SUIHAN Then
            sts = GetIni(App.EXEName, Trim(StrConv(OLD_Y_SYU2REC.HIN_NO, vbUnicode)), "SETUP", c)
            If Not sts Then
                Call UniCode_Conv(OLD_Y_SYU2REC.JGYOBU, SENTAKU)
                OUT_Count = OUT_Count + 1
            End If
        End If
            
        Do
            sts = BTRV(BtOpInsert, Y_SYU_POS, OLD_Y_SYU2REC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "出荷予定")
                    Exit Function
            End Select
        Loop
        

        com = BtOpGetNext
    
    Loop

    OUT_Cnt(3).Caption = Format(OUT_Count, "#0")    '炊飯⇒アイロン更新件数


'---------------------------------------------  終了
    Call Log_Out(LOG_F, "コンバート終了=" & Format(Now, "HH:MM:SS"))
    
    
    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "終了しました。"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
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
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）品目マスタＯＰＥＮ
    If OLD_ITEM2_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）在庫データＯＰＥＮ
    
    If OLD_ZAIKO2_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）在庫移動歴データＯＰＥＮ
    If OLD_IDO2_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）出荷予定データＯＰＥＮ
    If OLD_Y_SYU2_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
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
                                            '(旧)品目マスタCLOSE
    sts = BTRV(BtOpClose, OLD_ITEM2_POS, OLD_ITEM2REC, Len(OLD_ITEM2REC), K0_OLD_ITEM2, Len(K0_OLD_ITEM2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "（旧）品目マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '(旧)在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_ZAIKO2_POS, OLD_ZAIKO2REC, Len(OLD_ZAIKO2REC), K0_OLD_ZAIKO2, Len(K0_OLD_ZAIKO2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫データ")
        End If
    End If
    
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '(旧)在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_IDO2_POS, OLD_IDO2REC, Len(OLD_IDO2REC), K0_OLD_IDO2, Len(K0_OLD_IDO2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫移動歴")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
                                            '(旧)出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_Y_SYU2_POS, OLD_Y_SYU2REC, Len(OLD_Y_SYU2REC), K0_OLD_Y_SYU2, Len(K0_OLD_Y_SYU2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)出荷予定データ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_ONO1 = Nothing

    End
End Sub

