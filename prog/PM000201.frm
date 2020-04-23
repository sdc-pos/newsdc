VERSION 5.00
Begin VB.Form PM000201 
   Caption         =   "コードマスタメンテナンス"
   ClientHeight    =   10485
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12540
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
   ScaleHeight     =   10485
   ScaleWidth      =   12540
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   2
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   5340
      Index           =   0
      ItemData        =   "PM000201.frx":0000
      Left            =   240
      List            =   "PM000201.frx":0002
      TabIndex        =   6
      Top             =   3720
      Width           =   12015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   1
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   240
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      TabIndex        =   18
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "削 除"
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
      TabIndex        =   10
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      TabIndex        =   7
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblLIST_OP_NM2 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   10080
      TabIndex        =   29
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblLIST_OP_NM1 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   8640
      TabIndex        =   28
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "名　称"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   27
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "CODE"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "略称"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   25
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblBikou 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   4680
      TabIndex        =   24
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label lblOP_NM2 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblOP_NM1 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "名称"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "CODE"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   20
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "区分"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   19
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "PM000201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxC_CODE% = 0                'ｺｰﾄﾞ
Private Const ptxC_NAME% = 1                '名称
Private Const ptxC_RNAME% = 2               '略称

Private Const ptxOPTION1% = 3               'ｵﾌﾟｼｮﾝ1
Private Const ptxOPTION2% = 4               'ｵﾌﾟｼｮﾝ2

'コンボ用添字
Private Const pcmbDATA_KBN% = 0

'リスト用添字
Private Const plstP_CODE% = 0


Private W_Index As Integer


'Private Const LAST_UPDATE_DAY$ = "[PM00020] 2010.12.29 10:00"
Private Const LAST_UPDATE_DAY$ = "[PM00020] 2018.04.09 10:45"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000201.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000201)


    PM000201.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
    
    Error_Check_Proc = True
    
        
    Select Case Mode
    
        Case ptxC_CODE
    
            If Trim(Text1(Mode).Text) = "" Then
                MsgBox "入力した項目はエラーです。(CODE 必須入力)"
                Text1(Mode).SetFocus
                Exit Function
            End If
    
    
        Case ptxC_NAME          '2018.04.07
            If Trim(Text1(ptxC_RNAME).Text) = "" Then
                Text1(ptxC_RNAME).Text = Text1(ptxC_NAME).Text
            End If
    
    End Select
        
        
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    'ｺｰﾄﾞﾏｽﾀ読み込み
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, Right(Combo1(pcmbDATA_KBN).Text, 2))
    Call UniCode_Conv(K0_P_CODE.C_Code, CODE)
    
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
            'ﾚｺｰﾄﾞ内容の表示
                                            'ｺｰﾄﾞ
            Text1(ptxC_CODE).Text = StrConv(P_CODEREC.C_Code, vbUnicode)
                                            '名称
            Text1(ptxC_NAME).Text = StrConv(P_CODEREC.C_NAME, vbUnicode)
                                            '略称
            Text1(ptxC_RNAME).Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
                                            'ｵﾌﾟｼｮﾝ1
            Text1(ptxOPTION1).Text = StrConv(P_CODEREC.OPTION1, vbUnicode)
                                            'ｵﾌﾟｼｮﾝ2
            Text1(ptxOPTION2).Text = StrConv(P_CODEREC.OPTION2, vbUnicode)
        
        Case BtErrKeyNotFound
                                            'ｺｰﾄﾞ
'            Text1(ptxC_CODE).Text = ""
                                            '名称
            Text1(ptxC_NAME).Text = ""
                                            '略称
            Text1(ptxC_RNAME).Text = ""
                                            'ｵﾌﾟｼｮﾝ1
            Text1(ptxOPTION1).Text = ""
                                            'ｵﾌﾟｼｮﾝ2
            Text1(ptxOPTION2).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   リストボックス表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


    List_Disp_Proc = True
    
    List1(plstP_CODE).Clear
    
    'ｺｰﾄﾞﾏｽﾀ読み込み
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, Right(Combo1(pcmbDATA_KBN).Text, 2))
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    
    com = BtOpGetGreaterEqual
    
    
    Do
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Exit Function
        
        End Select

        If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> Right(Combo1(pcmbDATA_KBN).Text, 2) Then
            Exit Do
        
        End If
        
        List1(plstP_CODE).AddItem StrConv(P_CODEREC.C_Code, vbUnicode) & "   " & _
                                    StrConv(P_CODEREC.C_NAME, vbUnicode) & "  " & _
                                    StrConv(P_CODEREC.OPTION1, vbUnicode) & "  " & _
                                    StrConv(P_CODEREC.OPTION2, vbUnicode)
        com = BtOpGetNext
    
    Loop
        
    DoEvents

    If List1(plstP_CODE).ListCount = 0 Then
        
        W_Index = -1
        Text1(ptxC_CODE).SetFocus
    
    Else
    
        List1(plstP_CODE).SetFocus
        List1(plstP_CODE).ListIndex = 0
            
    End If

    List_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   コードマスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    'コードマスタ　読み込み
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, Right(Combo1(pcmbDATA_KBN).Text, 2))
    Call UniCode_Conv(K0_P_CODE.C_Code, Text1(ptxC_CODE).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CODE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "コードマスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    
    Call UniCode_Conv(P_CODEREC.DATA_KBN, Right(Combo1(pcmbDATA_KBN).Text, 2))          'ﾃﾞｰﾀ区分
    Call UniCode_Conv(P_CODEREC.C_Code, Text1(ptxC_CODE).Text)                          'コード
    Call UniCode_Conv(P_CODEREC.C_NAME, Text1(ptxC_NAME).Text)                          '名称
    Call UniCode_Conv(P_CODEREC.C_RNAME, Text1(ptxC_RNAME).Text)                        '略称
    Call UniCode_Conv(P_CODEREC.OPTION1, Text1(ptxOPTION1).Text)                        'オプション１
    Call UniCode_Conv(P_CODEREC.OPTION2, Text1(ptxOPTION2).Text)                        'オプション２
    
    
    Call UniCode_Conv(P_CODEREC.UPD_TANTO, "")                                          '更新担当者
                                                                                        '更新日時
    Call UniCode_Conv(P_CODEREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    Call UniCode_Conv(P_CODEREC.FILLER, "")                                             'Filler
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CODE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        End Select
    Loop
    
    For i = ptxC_CODE To ptxOPTION2
        Text1(i).Text = ""
    Next i


    If List_Disp_Proc() Then
        Exit Function
    End If
    
    List1(plstP_CODE).SetFocus
    If W_Index <> -1 Then
        List1(plstP_CODE).ListIndex = W_Index - 1
    End If
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   コードマスタ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

    Delete_Proc = True
    
    'コードマスタ　読み込み
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, Right(Combo1(pcmbDATA_KBN).Text, 2))
    Call UniCode_Conv(K0_P_CODE.C_Code, Text1(ptxC_CODE).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CODE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "コードマスタ")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CODE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "コードマスタ")
                Exit Function
        End Select
    Loop

    For i = ptxC_CODE To ptxOPTION1
        Text1(i).Text = ""
    Next i




    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstP_CODE).ListCount > 0 Then
        List1(plstP_CODE).SetFocus
        If W_Index - 1 > List1(plstP_CODE).ListCount Then
                List1(plstP_CODE).ListIndex = List1(plstP_CODE).ListCount
        Else
                
                List1(plstP_CODE).ListIndex = W_Index - 1
        End If
    Else
        Text1(ptxC_CODE).SetFocus
    End If


    Delete_Proc = False


End Function

Private Sub Combo1_Click(Index As Integer)
Dim i       As Integer      '2018.04.09
                
        
                                    'ｺｰﾄﾞ最大桁数設定
    Text1(ptxC_CODE).MaxLength = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_Len
                                    'ｵﾌﾟｼｮﾝ1使用有無
    Text1(ptxOPTION1).Visible = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP1
                                    'ｵﾌﾟｼｮﾝ2使用有無
    Text1(ptxOPTION2).Visible = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP2
                                    'ｵﾌﾟｼｮﾝﾀｲﾄﾙ1
    lblOP_NM1.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM1
                                    'ｵﾌﾟｼｮﾝﾀｲﾄﾙ2
    lblOP_NM2.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM2
                                    '備考
    lblBikou.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_BIKOU
    
    
                                    'ｵﾌﾟｼｮﾝﾀｲﾄﾙ1
    lblLIST_OP_NM1.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM1
                                    'ｵﾌﾟｼｮﾝﾀｲﾄﾙ2
    lblLIST_OP_NM2.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM2
    
    
    For i = ptxC_CODE To ptxOPTION2     '2018.04.09
        Text1(i).Text = ""              '2018.04.09
    Next i                              '2018.04.09
    
    
    
    
    If List_Disp_Proc() Then
        Unload Me
    End If
        
    Combo1(Index).SetFocus

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        
                                        'ｺｰﾄﾞ最大桁数設定
        Text1(ptxC_CODE).MaxLength = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_Len
                                        'ｵﾌﾟｼｮﾝ1使用有無
        Text1(ptxOPTION1).Visible = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP1
                                        'ｵﾌﾟｼｮﾝ2使用有無
        Text1(ptxOPTION2).Visible = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP2
                                        'ｵﾌﾟｼｮﾝﾀｲﾄﾙ1
        lblOP_NM1.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM1
                                        'ｵﾌﾟｼｮﾝﾀｲﾄﾙ2
        lblOP_NM2.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM2
                                        '備考
        lblBikou.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_BIKOU
        
        
                                        'ｵﾌﾟｼｮﾝﾀｲﾄﾙ1
        lblLIST_OP_NM1.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM1
                                        'ｵﾌﾟｼｮﾝﾀｲﾄﾙ2
        lblLIST_OP_NM2.Caption = P_KBN_TBL(Combo1(pcmbDATA_KBN).ListIndex).KBN_OP_NM2
        
        
        
        If List_Disp_Proc() Then
            Unload Me
        End If
        
        
'        Call Tab_Ctrl(Shift)        '移動
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
            For i = ptxC_CODE To ptxOPTION2
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
                    
        
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            End If
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm                  '2018.04.07
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim i       As Integer

'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If


    PM000201.Caption = PM000201.Caption & " " & LAST_UPDATE_DAY


'>  2018.04.07

                                'ログファイル名取り込み
'    If GetIni("FILE", "LOGF", "SYS", c) Then
'        Beep
'        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
    
    
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
'>  2018.04.07
    
    
    LOG_F = RTrim(c)
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
   
    Call P_CODE_INI_TBL_Proc    '2018.04.07
    
                                '区分設定
    Combo1(pcmbDATA_KBN).Clear
    For i = 0 To P_KBN_MAX
        Combo1(pcmbDATA_KBN).AddItem P_KBN_TBL(i).KBN_NM & "            " & P_KBN_TBL(i).KBN_CD
    Next i
                                
                                
                                
    W_Index = -1
    
    Show
    
    Combo1(pcmbDATA_KBN).ListIndex = 0
    
    If List_Disp_Proc() Then
        Unload Me
    End If
    
    Combo1(pcmbDATA_KBN).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000201 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)

Dim W_KEY   As String

    W_Index = List1(plstP_CODE).ListIndex
    W_KEY = Left(List1(plstP_CODE).List(List1(plstP_CODE).ListIndex), 10)

    
    If Item_Disp_Proc(W_KEY) Then     '明細表示
        Unload Me
    End If

End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(plstP_CODE).ListCount > 0 And _
       List1(plstP_CODE).ListIndex < 0 Then
        List1(plstP_CODE).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim W_KEY   As String
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        W_Index = List1(plstP_CODE).ListIndex
        W_KEY = Left(List1(plstP_CODE).List(List1(plstP_CODE).ListIndex), 10)
    
        
        If Item_Disp_Proc(W_KEY) Then     '明細表示
            Unload Me
        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

