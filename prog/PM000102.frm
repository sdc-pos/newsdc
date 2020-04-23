VERSION 5.00
Begin VB.Form PM000102 
   Caption         =   "管理マスタメンテナンス"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12045
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
   ScaleHeight     =   6300
   ScaleWidth      =   12045
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   15
      Text            =   "99999999"
      Top             =   1320
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   14
      Text            =   "99999999"
      Top             =   840
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   13
      Text            =   "99999999"
      Top             =   360
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "999"
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "999"
      Top             =   1920
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "999"
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "9999.99"
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "9999.99"
      Top             =   840
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "999.99"
      Top             =   480
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "9999.99"
      Top             =   2640
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   5
      Text            =   "9999.99"
      Top             =   2280
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "9999.99"
      Top             =   1800
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "9999.99"
      Top             =   1440
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "9999.99"
      Top             =   960
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "9999.99"
      Top             =   600
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   27
      Top             =   5880
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "次画面"
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   3
      Left            =   2640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "売上№"
      Height          =   240
      Index           =   15
      Left            =   7980
      TabIndex        =   43
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "請求書№"
      Height          =   240
      Index           =   14
      Left            =   7770
      TabIndex        =   42
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "見積書№"
      Height          =   240
      Index           =   13
      Left            =   7770
      TabIndex        =   41
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　ラベル貼付枚数"
      Height          =   240
      Index           =   12
      Left            =   3780
      TabIndex        =   40
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　同梱品点検件数"
      Height          =   240
      Index           =   11
      Left            =   3780
      TabIndex        =   39
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　副資材点検件数"
      Height          =   240
      Index           =   10
      Left            =   3780
      TabIndex        =   38
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　　余裕率"
      Height          =   240
      Index           =   9
      Left            =   3780
      TabIndex        =   37
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　　分レート"
      Height          =   240
      Index           =   8
      Left            =   3780
      TabIndex        =   36
      Top             =   960
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "工程　　ロット数"
      Height          =   240
      Index           =   7
      Left            =   3780
      TabIndex        =   35
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　余裕率"
      Height          =   240
      Index           =   6
      Left            =   630
      TabIndex        =   34
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　余裕率"
      Height          =   240
      Index           =   5
      Left            =   630
      TabIndex        =   33
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "出荷　分レート"
      Height          =   240
      Index           =   4
      Left            =   630
      TabIndex        =   32
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "出庫　分レート"
      Height          =   240
      Index           =   3
      Left            =   630
      TabIndex        =   31
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "　　　余裕率"
      Height          =   240
      Index           =   1
      Left            =   630
      TabIndex        =   30
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "入庫　分レート"
      Height          =   240
      Index           =   0
      Left            =   630
      TabIndex        =   29
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label Label 
      Caption         =   "ﾚｺｰﾄﾞ№"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PM000102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxRec_No% = 0                'ﾚｺｰﾄﾞ№(入力不可)


Private Const ptxNYUKO_S_RATE% = 1          '入庫　分レート
Private Const ptxNYUKO_R_RATE% = 2          '入庫　余裕率

Private Const ptxSYUKO_S_RATE% = 3          '出庫　分レート
Private Const ptxSYUKO_R_RATE% = 4          '出庫　余裕率

Private Const ptxSYUKA_S_RATE% = 5          '出荷　分レート
Private Const ptxSYUKA_R_RATE% = 6          '出荷　余裕率

Private Const ptxKOUTEI_LOT% = 7            '工程　前後工程標準ロット
Private Const ptxKOUTEI_S_RATE% = 8         '工程　分レート
Private Const ptxKOUTEI_R_RATE% = 9         '工程　余裕率
Private Const ptxKOUTEI_SHIZAI% = 10        '工程　副資材確認点数
Private Const ptxKOUTEI_BUHIN% = 11         '工程　同梱部品確認点数
Private Const ptxKOUTEI_LABEL% = 12         '工程　ラベル貼付枚数

Private Const ptxMITSUMORI_NO% = 13         '見積書№
Private Const ptxSEIKYU_NO% = 14            '請求書№
Private Const ptxMIN_URIAGE_NO% = 15        'ミニマム　売上№







Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000102.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000102)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000102)


    PM000102.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxNYUKO_S_RATE            '入庫　分レート

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxNYUKO_R_RATE            '入庫　余裕率
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKO_S_RATE            '出庫　分レート

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKO_R_RATE                '出庫　余裕率

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKA_S_RATE                '出荷　分レート

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKA_R_RATE                '出荷　余裕率
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_LOT                  '工程　前後工程標準ロット
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxKOUTEI_S_RATE               '工程　分レート

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_R_RATE               '工程　余裕率

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_SHIZAI               '工程　副資材確認点数

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_BUHIN                '工程　同梱部品確認点数
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxKOUTEI_LABEL                '工程　ラベル貼付枚数
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxMITSUMORI_NO                '見積書№
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            End If
        Case ptxSEIKYU_NO                   '請求書№
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            End If
        Case ptxMIN_URIAGE_NO               'ミニマム売上№
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            End If
    
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '管理ﾏｽﾀ（KEY=0）読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    
    
    Text1(ptxRec_No).Text = P_ST_KANRI_No   'ﾚｺｰﾄﾞ№
    
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            'ﾚｺｰﾄﾞ内容の表示
            
            If IsNumeric(StrConv(P_KANRIREC.NYUKO_S_RATE, vbUnicode)) Then
                Text1(ptxNYUKO_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.NYUKO_S_RATE, vbUnicode)), "#0.00")    '入庫　分レート
            Else
                Text1(ptxNYUKO_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.NYUKO_R_RATE, vbUnicode)) Then
                Text1(ptxNYUKO_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.NYUKO_R_RATE, vbUnicode)), "#0.00")    '入庫　余裕率
            Else
                Text1(ptxNYUKO_R_RATE).Text = ""
            End If
            
            If IsNumeric(StrConv(P_KANRIREC.SYUKO_S_RATE, vbUnicode)) Then
                Text1(ptxSYUKO_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKO_S_RATE, vbUnicode)), "#0.00")    '出庫　分レート
            Else
                Text1(ptxSYUKO_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKO_R_RATE, vbUnicode)) Then
                Text1(ptxSYUKO_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKO_R_RATE, vbUnicode)), "#0.00")    '出庫　余裕率
            Else
                Text1(ptxSYUKO_R_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKA_S_RATE, vbUnicode)) Then
                Text1(ptxSYUKA_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKA_S_RATE, vbUnicode)), "#0.00")    '出荷　分レート
            Else
                Text1(ptxSYUKA_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKA_R_RATE, vbUnicode)) Then
                Text1(ptxSYUKA_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKA_R_RATE, vbUnicode)), "#0.00")    '出荷　分レート
            Else
                Text1(ptxSYUKA_R_RATE).Text = ""
            End If
            
            
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
                Text1(ptxKOUTEI_LOT).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#0.00")        '工程　前後工程標準ロット
            Else
                Text1(ptxKOUTEI_LOT).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
                Text1(ptxKOUTEI_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")  '工程　分レート
            Else
                Text1(ptxKOUTEI_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
                Text1(ptxKOUTEI_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")  '工程　余裕率
            Else
                Text1(ptxKOUTEI_R_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_SHIZAI, vbUnicode)) Then
                Text1(ptxKOUTEI_SHIZAI).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_SHIZAI, vbUnicode)), "#")      '工程　副資材確認点数
            Else
                Text1(ptxKOUTEI_SHIZAI).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_BUHIN, vbUnicode)) Then
                Text1(ptxKOUTEI_BUHIN).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_BUHIN, vbUnicode)), "#")       '工程　同梱部品確認点数
            Else
                Text1(ptxKOUTEI_BUHIN).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LABEL, vbUnicode)) Then
                Text1(ptxKOUTEI_LABEL).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_LABEL, vbUnicode)), "#")       '工程　ラベル貼付け枚数
            Else
                Text1(ptxKOUTEI_LABEL).Text = ""
            End If
  
            If IsNumeric(StrConv(P_KANRIREC.MITSUMORI_NO, vbUnicode)) Then
                Text1(ptxMITSUMORI_NO).Text = Format(CDbl(StrConv(P_KANRIREC.MITSUMORI_NO, vbUnicode)), "00000000")     '見積書№
            Else
                Text1(ptxMITSUMORI_NO).Text = "00000001"
            End If
            
            If IsNumeric(StrConv(P_KANRIREC.SEIKYU_NO, vbUnicode)) Then
                Text1(ptxSEIKYU_NO).Text = Format(CDbl(StrConv(P_KANRIREC.SEIKYU_NO, vbUnicode)), "00000000")           '請求書№
            Else
                Text1(ptxSEIKYU_NO).Text = "00000001"
            End If
            If IsNumeric(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)) Then
                Text1(ptxMIN_URIAGE_NO).Text = Format(CDbl(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)), "00000000")   'ミニマム売上№
            Else
                Text1(ptxMIN_URIAGE_NO).Text = "00000001"
            End If
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   管理マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer


    Update_Proc = True
    '管理ﾏｽﾀ（KEY=0）読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    
    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)                                 'ﾚｺｰﾄﾞ№
    
    If IsNumeric(Text1(ptxNYUKO_S_RATE).Text) Then                                      '入庫　分レート
        Call UniCode_Conv(P_KANRIREC.NYUKO_S_RATE, Format(CDbl(Text1(ptxNYUKO_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.NYUKO_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxNYUKO_R_RATE).Text) Then                                      '入庫　余裕率
        Call UniCode_Conv(P_KANRIREC.NYUKO_R_RATE, Format(CDbl(Text1(ptxNYUKO_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.NYUKO_R_RATE, "")
    End If
    
    
    If IsNumeric(Text1(ptxSYUKO_S_RATE).Text) Then                                      '出庫　分レート
        Call UniCode_Conv(P_KANRIREC.SYUKO_S_RATE, Format(CDbl(Text1(ptxSYUKO_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKO_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxSYUKO_R_RATE).Text) Then                                      '出庫　余裕率
        Call UniCode_Conv(P_KANRIREC.SYUKO_R_RATE, Format(CDbl(Text1(ptxSYUKO_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKO_R_RATE, "")
    End If
    
    If IsNumeric(Text1(ptxSYUKA_S_RATE).Text) Then                                      '出荷　分レート
        Call UniCode_Conv(P_KANRIREC.SYUKA_S_RATE, Format(CDbl(Text1(ptxSYUKA_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKA_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxSYUKA_S_RATE).Text) Then                                      '出荷　余裕率
        Call UniCode_Conv(P_KANRIREC.SYUKA_R_RATE, Format(CDbl(Text1(ptxSYUKA_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKA_R_RATE, "")
    End If
    
    If IsNumeric(Text1(ptxKOUTEI_LOT).Text) Then                                        '工程　前後工程標準ロット
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LOT, Format(CDbl(Text1(ptxKOUTEI_LOT).Text), "000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LOT, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_S_RATE).Text) Then                                     '工程　分レート
        Call UniCode_Conv(P_KANRIREC.KOUTEI_S_RATE, Format(CDbl(Text1(ptxKOUTEI_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_R_RATE).Text) Then                                     '工程　余裕率
        Call UniCode_Conv(P_KANRIREC.KOUTEI_R_RATE, Format(CDbl(Text1(ptxKOUTEI_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_R_RATE, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_SHIZAI).Text) Then                                     '工程　副資材確認点数
        Call UniCode_Conv(P_KANRIREC.KOUTEI_SHIZAI, Format(CDbl(Text1(ptxKOUTEI_SHIZAI).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_SHIZAI, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_BUHIN).Text) Then                                      '工程　同梱部品確認点数
        Call UniCode_Conv(P_KANRIREC.KOUTEI_BUHIN, Format(CDbl(Text1(ptxKOUTEI_BUHIN).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_BUHIN, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_LABEL).Text) Then                                      '工程　ラベル貼付枚数
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LABEL, Format(CDbl(Text1(ptxKOUTEI_LABEL).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LABEL, "")
    End If
                                                                                        '見積書№
    Call UniCode_Conv(P_KANRIREC.MITSUMORI_NO, Format(CDbl(Text1(ptxMITSUMORI_NO).Text), "00000000"))
                                                                                        '請求書№
    Call UniCode_Conv(P_KANRIREC.SEIKYU_NO, Format(CDbl(Text1(ptxSEIKYU_NO).Text), "00000000"))
                                                                                        'ミニマム売上№
    Call UniCode_Conv(P_KANRIREC.MIN_URIAGE_NO, Format(CDbl(Text1(ptxMIN_URIAGE_NO).Text), "00000000"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "管理マスタ")
                Exit Function
        End Select
    Loop

    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            For i = ptxRec_No To ptxSEIKYU_NO
            
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
            Text1(ptxNYUKO_S_RATE).SetFocus
        Case P_CMD_DEL                      '削除
        Case P_CMD_DSP                      '検索/表示
        
        Case 5                              '初期値設定   2008.02.13
        
            PM000103.Show vbModal
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Me.Visible = False
    End Select

End Sub


Private Sub Form_Activate()
                                '画面初期設定
    If Item_Disp_Proc() Then
        Unload Me
    End If
                                
    Text1(ptxNYUKO_S_RATE).SetFocus

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
    PM000102.Caption = PM000102.Caption & LAST_UPDATE_DAY

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

