VERSION 5.00
Begin VB.Form PM000702 
   Caption         =   "受払先マスタメンテナンス"
   ClientHeight    =   7155
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12645
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
   ScaleHeight     =   7155
   ScaleWidth      =   12645
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PM000702.frx":0000
      Left            =   2160
      List            =   "PM000702.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   9
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   11
      Top             =   5520
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   8
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   10
      Top             =   5040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   4
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   3
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   2
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1920
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      ItemData        =   "PM000702.frx":0004
      Left            =   2640
      List            =   "PM000702.frx":0006
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   840
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "戻 る"
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
      TabIndex        =   23
      Top             =   6240
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   15
      Top             =   6240
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "取引先区分"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   36
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "郵便番号"
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   35
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "住所２"
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   34
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "住所１"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   33
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "（「-」含む）"
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   32
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "FAX番号"
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   31
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "（「-」含む）"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "電話番号"
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "部署／営業所名"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "略称"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   27
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "受払先名称"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "収支ｺｰﾄﾞ"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "受払先ｺｰﾄﾞ"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "PM000702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxUKEHARAI_CODE% = 0         '受払い先ｺｰﾄﾞ
Private Const ptxSYUSHI_CODE% = 1           '収支ｺｰﾄﾞ
Private Const ptxUKEHARAI_NAME% = 2         '受払先名称
Private Const ptxUKEHARAI_RNAME% = 3        '受払先略称
Private Const ptxBUSHO_NAME% = 4            '部署名称
Private Const ptxTEL_NO% = 5                '電話番号
Private Const ptxFAX_NO% = 6                'FAX番号
Private Const ptxYUBIN_NO% = 7              '郵便番号
Private Const ptxADDR1% = 8                 '住所１
Private Const ptxADDR2% = 9                 '住所２

Private Const Mode_All% = 0
'コンボ用添字
Private Const pcmbSYUSHI% = 0               '収支
Private Const pcmbTORI_KBN% = 1             '取引先区分



Private INIT_FLG    As Boolean
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000701.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000701)


    PM000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer
    
Dim i       As Integer          '2019.04.04
    
    Error_Check_Proc = True
    
    
    
    Select Case Mode
        
        Case Mode_All, ptxUKEHARAI_CODE     '受払先コード
            
            Text1(Mode).Text = StrConv(RTrim(Text1(Mode).Text), vbUpperCase)
            
            
            
            If Trim(Text1(ptxUKEHARAI_CODE).Text) = "" Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxUKEHARAI_CODE).SetFocus
                Exit Function
            End If
            
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                '新規時は重複チェック
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
            
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                        ans = MsgBox("入力したコードは、登録済です。更新処理として継続しますか？")
                        If ans = vbNo Then
                            Text1(ptxUKEHARAI_CODE).SetFocus
                            Exit Function
                        End If
                    
                        Call Item_Disp_Proc(Text1(ptxUKEHARAI_CODE).Text)
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                        Exit Function
                End Select
            
            
                Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_NG
                Text1(ptxUKEHARAI_CODE).Locked = True
                Text1(ptxUKEHARAI_CODE).TabStop = False
            
            End If
        
        Case Mode_All, ptxSYUSHI_CODE       '収支ｺｰﾄﾞ
        
'>>>>>>>>>>>>>> 2019.04.04
            For i = 0 To Combo1(pcmbSYUSHI).ListCount - 1
                    
                If Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).List(i), 3) Then
                    Combo1(pcmbSYUSHI).ListIndex = i
                    Exit For
                End If
            
            Next i
        
            If i > (Combo1(pcmbSYUSHI).ListCount - 1) Then
                MsgBox ("入力した収支コードは、未登録です。再入力して下さい。")
                Text1(ptxSYUSHI_CODE).SetFocus
                Exit Function
            End If
'>>>>>>>>>>>>>> 2019.04.04
        
        
        
        
        
        
        Case Mode_All, ptxUKEHARAI_NAME     '受払先名称
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                Text1(ptxUKEHARAI_RNAME).Text = Text1(ptxUKEHARAI_NAME).Text
                    
            End If
        
        
        Case Mode_All, ptxUKEHARAI_RNAME    '受払先略称
        Case Mode_All, ptxBUSHO_NAME        '部署／営業所
        Case Mode_All, ptxTEL_NO            '電話番号
        Case Mode_All, ptxFAX_NO            'FAX番号
        Case Mode_All, ptxYUBIN_NO          '郵便番号
        Case Mode_All, ptxADDR1             '住所１
        Case Mode_All, ptxADDR2             '住所２
              
        
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
    
    '受払先ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, CODE)
    
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
            'ﾚｺｰﾄﾞ内容の表示
                                            '受払先ｺｰﾄﾞ
            Text1(ptxUKEHARAI_CODE).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode))
                                            '収支ｺｰﾄﾞ
            Text1(ptxSYUSHI_CODE).Text = Trim(StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode))
                                            '収支名検索
            For i = 0 To Combo1(pcmbSYUSHI).ListCount - 1
                If Right(Combo1(pcmbSYUSHI).List(i), 3) = Text1(ptxSYUSHI_CODE).Text Then
                    Combo1(pcmbSYUSHI).ListIndex = i
                    Exit For
                End If
            
            Next i
                                            '受払先名称
            Text1(ptxUKEHARAI_NAME).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
                                            '受払先略称
            Text1(ptxUKEHARAI_RNAME).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
                                            '部署／営業所名
            Text1(ptxBUSHO_NAME).Text = Trim(StrConv(P_UKEHARAIREC.BUSHO_NAME, vbUnicode))
                                            '電話番号
            Text1(ptxTEL_NO).Text = Trim(StrConv(P_UKEHARAIREC.TEL_NO, vbUnicode))
                                            'FAX番号
            Text1(ptxFAX_NO).Text = Trim(StrConv(P_UKEHARAIREC.FAX_NO, vbUnicode))
                                            '郵便番号
            Text1(ptxYUBIN_NO).Text = Trim(StrConv(P_UKEHARAIREC.YUBIN_NO, vbUnicode))
                                            '住所１
            Text1(ptxADDR1).Text = Trim(StrConv(P_UKEHARAIREC.ADDR1, vbUnicode))
                                            '住所２
            Text1(ptxADDR2).Text = Trim(StrConv(P_UKEHARAIREC.ADDR2, vbUnicode))
                                            '取引先区分
            Combo1(pcmbTORI_KBN).ListIndex = 0
            For i = 0 To Combo1(pcmbTORI_KBN).ListCount - 1
                If Right(Combo1(pcmbTORI_KBN).List(i), 1) = StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) Then
                    Combo1(pcmbTORI_KBN).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        
        Case BtErrKeyNotFound
        
            MsgBox "他端末で変更されています。前画面に戻ります。"
            PM000702.Visible = False
            INIT_FLG = False
            
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            PM000702.Visible = False
            INIT_FLG = False
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   受払先マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    '受払先マスタ　読み込み
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "受払先マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)        '受払先ｺｰﾄﾞ
    Call UniCode_Conv(P_UKEHARAIREC.SYUSHI_CODE, Text1(ptxSYUSHI_CODE).Text)            '収支ｺｰﾄﾞ
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, Text1(ptxUKEHARAI_NAME).Text)        '受払先名
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, Text1(ptxUKEHARAI_RNAME).Text)      '受払略称
    Call UniCode_Conv(P_UKEHARAIREC.BUSHO_NAME, Text1(ptxBUSHO_NAME).Text)              '部署／営業所
    Call UniCode_Conv(P_UKEHARAIREC.TEL_NO, Text1(ptxTEL_NO).Text)                      '電話番号
    Call UniCode_Conv(P_UKEHARAIREC.FAX_NO, Text1(ptxFAX_NO).Text)                      'FAX番号
    Call UniCode_Conv(P_UKEHARAIREC.YUBIN_NO, Text1(ptxYUBIN_NO).Text)                  '郵便番号
    Call UniCode_Conv(P_UKEHARAIREC.ADDR1, Text1(ptxADDR1).Text)                        '住所１
    Call UniCode_Conv(P_UKEHARAIREC.ADDR2, Text1(ptxADDR2).Text)                        '住所２
    Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, Right(Combo1(pcmbTORI_KBN).Text, 1))      '取引先区分
    
    
    
    
    Call UniCode_Conv(P_UKEHARAIREC.FILLER, "")                                         'Filler
    
    Call UniCode_Conv(P_UKEHARAIREC.UPD_TANTO, "")                                      '更新担当者ｺｰﾄﾞ
                                                                                        '更新日時
    Call UniCode_Conv(P_UKEHARAIREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        End Select
    
    Loop
    
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   受払先マスタ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '受払先マスタ　読み込み
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "受払先マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_UKEHRAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "受払先マスタ")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbSYUSHI     '収支
            Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).Text, 3)
    
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Command1_Click(Index As Integer)

Dim yn As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
            If Error_Check_Proc(0) Then     'エラーチェック
                Exit Sub
            End If
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    PM000702.Visible = False
                    INIT_FLG = False
                End If
            End If
            PM000702.Visible = False
            INIT_FLG = False
                    
        
        
        Case P_CMD_DEL                      '削除
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Delete_Proc() Then
                    PM000702.Visible = False
                    INIT_FLG = False
                End If
            End If
            PM000702.Visible = False
            INIT_FLG = False
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            PM000702.Visible = False
            INIT_FLG = False
    End Select

End Sub

Private Sub Form_Activate()
    
Dim i       As Integer
Dim CODE    As String
    
    If INIT_FLG Then
        Exit Sub
    End If


    Select Case G_SCREEN_FLG
        Case G_SCREEN_INS       '新規
                
            Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_OK
            Text1(ptxUKEHARAI_CODE).TabStop = True
            Text1(ptxUKEHARAI_CODE).Locked = False
                
            For i = ptxUKEHARAI_CODE To ptxADDR2
                Text1(i).Text = ""
            Next i
            
            Combo1(pcmbTORI_KBN).ListIndex = 0
                
            Text1(ptxUKEHARAI_CODE).SetFocus
                
        
        Case G_SCREEN_UPD       '更新
    
                
    
    
            Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_NG
            Text1(ptxUKEHARAI_CODE).TabStop = False
            Text1(ptxUKEHARAI_CODE).Locked = True
    
            
            CODE = PM000701.txSEL_KEY
            
            If Item_Disp_Proc(CODE) Then
                Unload Me
            End If
    
            Text1(ptxSYUSHI_CODE).SetFocus
    
    End Select


    INIT_FLG = True

End Sub

Private Sub Form_DblClick()
'    PrintForm
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

Dim com     As Integer
Dim sts     As Integer

    
    PM000702.Caption = PM000702.Caption & LAST_UPDATE_DAY
    
    
    '収支内容のセット
    If Code_Set_Proc(pcmbSYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    '取引先区分
    Combo1(pcmbTORI_KBN).Clear
    Combo1(pcmbTORI_KBN).AddItem P_TORI_GENERAL_N & "    " & P_TORI_GENERAL
    Combo1(pcmbTORI_KBN).AddItem P_TORI_NAISYOKU_N & "    " & P_TORI_NAISYOKU
    Combo1(pcmbTORI_KBN).AddItem P_TORI_GENKIN_N & "    " & P_TORI_GENKIN
    Combo1(pcmbTORI_KBN).AddItem P_TORI_SYANAI_N & "    " & P_TORI_SYANAI
    Combo1(pcmbTORI_KBN).AddItem P_TORI_ANOTHER_N & "    " & P_TORI_ANOTHER
    
    Combo1(pcmbTORI_KBN).AddItem P_TORI_JIKYU_N & "    " & P_TORI_JIKYU
    
    
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            
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
    Set PM000701 = Nothing
    Set PM000702 = Nothing

    End
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    End Select
        
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    




End Function


