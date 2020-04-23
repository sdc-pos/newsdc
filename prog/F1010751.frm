VERSION 5.00
Begin VB.Form F1010751 
   BackColor       =   &H00C0C0C0&
   Caption         =   "メニュー管理マスタセットアップ"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "メニューマスタ更新中！"
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
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "メニュー管理マスタセットアップ"
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
      Visible         =   0   'False
      Width           =   7200
   End
End
Attribute VB_Name = "F1010751"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GLB_MENU_NO         As String * 2       '共通メニュー番号
Private GLB_MENU_NAME       As String           '共通メニュー名称


Private NAIGAI_CODE()       As String * 1       '内外テーブル


Private Type YOIN_TBL_Tag                       '要因テーブル（メニューの先頭）
    CODE_TYPE               As String * 1       '主バーコードタイプ
    CODE_NAME               As String * 5       '主バーコード名称
End Type

Private YOIN_TBL()          As YOIN_TBL_Tag
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  最新共通メニュー作成処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com_MENU    As Integer
Dim com_YOIN    As Integer
Dim com_MTS     As Integer
Dim ans         As Integer

Dim LEVEL_NO1   As Integer
Dim LEVEL_NO2   As Integer
Dim LEVEL_NO3   As Integer
        
Dim i           As Integer
Dim j           As Integer
    
    Update_Proc = True
    
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
    Me.MousePointer = vbHourglass
    
    '----------------------------   対象メニュー全件削除
    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, GLB_MENU_NO)
    Call UniCode_Conv(K0_MENU.JGYOBU, "")
    Call UniCode_Conv(K0_MENU.NAIGAI, "")
    Call UniCode_Conv(K0_MENU.MENU_LV1, "")
    Call UniCode_Conv(K0_MENU.MENU_LV2, "")
    Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
    com_MENU = BtOpGetGreaterEqual
    
    Do
        DoEvents
        Do
            sts = BTRV(com_MENU + BtSNoWait, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.MENU_GRP_NO, vbUnicode) <> GLB_MENU_NO Then
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com_MENU + BtSNoWait, "メニュー管理")
                    Exit Function
            End Select
        
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
        
            sts = BTRV(BtOpDelete, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "メニュー管理")
                    Exit Function
            End Select
        
        
        Loop
        
        com_MENU = BtOpGetNext
    
    Loop

    '----------------------------   メニュー作成開始
    For i = 0 To UBound(JGYOBU_T)                   '事業部のループ
        
        For j = 0 To UBound(NAIGAI_CODE)                '国内外のループ
            
            For LEVEL_NO1 = 0 To UBound(YOIN_TBL)           '要因のループ

                Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
                Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
                Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
                Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
                Call UniCode_Conv(MENUREC.MENU_LV2, "")
                Call UniCode_Conv(MENUREC.MENU_LV3, "")
                
                Call UniCode_Conv(MENUREC.MENU_KBN, "0")
                Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
                Call UniCode_Conv(MENUREC.DISPLAY_ITEM, YOIN_TBL(LEVEL_NO1).CODE_NAME)

                Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                Call UniCode_Conv(MENUREC.YOIN_CODE, "")
                Call UniCode_Conv(MENUREC.PARAM, "")
                Call UniCode_Conv(MENUREC.FILLER, "")
            
                Do
                    sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                
                        Case Else
                            Call File_Error(sts, BtOpDelete, "メニュー管理")
                            Exit Function
                    End Select
                Loop
                '------------------------   該当要因マスタＳＴＡＲＴ
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
                        
                LEVEL_NO2 = 0
                com_YOIN = BtOpGetGreater
                Do
                    DoEvents
                    sts = BTRV(com_YOIN, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> YOIN_TBL(LEVEL_NO1).CODE_TYPE Then
                                Exit Do
                            End If
                        Case BtErrEOF
                            Exit Do
                
                        Case Else
                            Call File_Error(sts, com_YOIN, "要因マスタ")
                            Exit Function
                    End Select
                
                
                    If StrConv(YOINREC.REGI_F, vbUnicode) = "0" Or StrConv(YOINREC.REGI_F, vbUnicode) = "1" Then
                
                        Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
                        Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
                        Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
                        Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
                        Call UniCode_Conv(MENUREC.MENU_LV2, Format(LEVEL_NO2, "000"))
                        Call UniCode_Conv(MENUREC.MENU_LV3, "")
                        
                        Call UniCode_Conv(MENUREC.MENU_KBN, "1")
                        Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
                        Call UniCode_Conv(MENUREC.DISPLAY_ITEM, StrConv(YOINREC.YOIN_DNAME, vbUnicode))
    
                        Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
                        Call UniCode_Conv(MENUREC.YOIN_CODE, StrConv(YOINREC.YOIN_CODE, vbUnicode))
                        Call UniCode_Conv(MENUREC.PARAM, "")
                        Call UniCode_Conv(MENUREC.FILLER, "")
                
                        Do
                            sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                    
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "メニュー管理")
                                    Exit Function
                            End Select
                        Loop
                    
'                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then
'
'                        '------------------------------ 向け先ＳＴＡＲＴ
'                            com_MTS = BtOpGetFirst
'
'                            LEVEL_NO3 = 0
'                            Do
'                                DoEvents
'                                sts = BTRV(com_MTS, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                    Case BtErrEOF
'                                        Exit Do
'
'                                    Case Else
'                                        Call File_Error(sts, com_MTS, "向け先管理マスタ")
'                                        Exit Function
'                                End Select
'
'
'                                Call UniCode_Conv(MENUREC.MENU_GRP_NO, GLB_MENU_NO)
'                                Call UniCode_Conv(MENUREC.JGYOBU, JGYOBU_T(i).Code)
'                                Call UniCode_Conv(MENUREC.NAIGAI, NAIGAI_CODE(j))
'                                Call UniCode_Conv(MENUREC.MENU_LV1, Format(LEVEL_NO1, "000"))
'                                Call UniCode_Conv(MENUREC.MENU_LV2, Format(LEVEL_NO2, "000"))
'                                Call UniCode_Conv(MENUREC.MENU_LV3, Format(LEVEL_NO3, "000"))
'
'                                Call UniCode_Conv(MENUREC.MENU_KBN, "1")
'                                Call UniCode_Conv(MENUREC.MENU_GRP, GLB_MENU_NAME)
'                                Call UniCode_Conv(MENUREC.DISPLAY_ITEM, StrConv(MTSREC.MUKE_DNAME, vbUnicode))
'
'                                Call UniCode_Conv(MENUREC.CODE_TYPE, YOIN_TBL(LEVEL_NO1).CODE_TYPE)
'                                Call UniCode_Conv(MENUREC.YOIN_CODE, StrConv(YOINREC.YOIN_CODE, vbUnicode))
'                                Call UniCode_Conv(MENUREC.PARAM, StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode))
'                                Call UniCode_Conv(MENUREC.FILLER, "")
'
'                                Do
'                                    sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
'                                    Select Case sts
'                                        Case BtNoErr
'                                            Exit Do
'                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
'
'                                        Case Else
'                                            Call File_Error(sts, BtOpDelete, "メニュー管理")
'                                        Exit Function
'                                    End Select
'                                Loop
'
'
'
'                                LEVEL_NO3 = LEVEL_NO3 + 1
'
'                                com_MTS = BtOpGetNext
'
'
'                            Loop
'
'
'
'                        End If
                    
                    
                        LEVEL_NO2 = LEVEL_NO2 + 1
                    
                    End If
                    
                    com_YOIN = BtOpGetNext
                
                
                
                Loop
            
            
            
            
            Next LEVEL_NO1                                  '要因のループ

        Next j                                      '国内外のループ
    
    Next i                                      '事業部のループ

    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        If Update_Proc() Then
            MsgBox "異常終了しました。"
            Unload Me
        End If
    End If
    MsgBox "正常終了しました。"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer
Dim CODE_TYPE   As String * 1
    
    
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
                                
                                '事業部の獲得
    If JGYOB_TB_Set() Then
        Beep
        MsgBox "事業部の獲得に失敗しました。"
        End
    End If
                                '国内外管理の獲得
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        Beep
        MsgBox "国内外の獲得に失敗しました。"
        End
    End If
                                '共通メニュー番号取り込み
    If GetIni(App.EXEName, "GLB_MENU_NO", "SYS", c) Then
        Beep
        MsgBox "共通メニュー番号の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    GLB_MENU_NO = RTrim(c)
                                
                                '共通メニュー名称取り込み
    If GetIni(App.EXEName, "GLB_MENU_NM", "SYS", c) Then
        Beep
        MsgBox "共通メニュー番号の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    GLB_MENU_NAME = RTrim(c)
                                
                                '要因情報取り込み
    i = -1
    j = 1
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(j, "00"), "SYS", c) Then
            Beep
            MsgBox "要因情報[ACTION_CD]の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
                
        If Trim(c) = "NON" Then
            Exit Do
        End If
        CODE_TYPE = Trim(c)
           
    
        If GetIni("ACTION", "ACTION_TYPE" & Format(j, "00"), "SYS", c) Then
            Beep
            MsgBox "要因情報[ACTION_TYPE]の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
    
        If Trim(c) = "1" Then
            'メニュー登録不可
        Else
            'メニュー登録可
            
            i = i + 1
            
            ReDim Preserve YOIN_TBL(i)
            YOIN_TBL(i).CODE_TYPE = CODE_TYPE
        
            If GetIni("ACTION", "ACTION_NM" & Format(j, "00"), "SYS", c) Then
                Beep
                MsgBox "要因情報[ACTION_NM]の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
            YOIN_TBL(i).CODE_NAME = Trim(c)
        End If
    
        j = j + 1
    Loop
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'メニュー管理マスタＯＰＥＮ
    If MENU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            'メニュー管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010751 = Nothing

    End
End Sub

