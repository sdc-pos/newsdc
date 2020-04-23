VERSION 5.00
Begin VB.Form F1030301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "標準棚番別出庫表印刷"
   ClientHeight    =   6945
   ClientLeft      =   2325
   ClientTop       =   2715
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
   ScaleHeight     =   6945
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷中止"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2520
      Width           =   615
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷先"
      Height          =   252
      Index           =   10
      Left            =   2520
      TabIndex        =   36
      Top             =   1920
      Width           =   732
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
      TabIndex        =   35
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   252
      Index           =   9
      Left            =   5760
      TabIndex        =   34
      Top             =   3360
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標準棚番"
      Height          =   252
      Index           =   8
      Left            =   2520
      TabIndex        =   33
      Top             =   3360
      Width           =   972
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   7
      Left            =   7200
      TabIndex        =   32
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   6
      Left            =   6720
      TabIndex        =   31
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   5
      Left            =   5760
      TabIndex        =   30
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   2
      Left            =   4920
      TabIndex        =   29
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   1
      Left            =   4440
      TabIndex        =   28
      Top             =   2640
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   27
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷区分"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   26
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷予定日"
      Height          =   252
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_DEN_DT_YY% = 0           '開始　出荷予定日　年
Private Const ptxS_DEN_DT_MM% = 1           '開始　出荷予定日　月
Private Const ptxS_DEN_DT_DD% = 2           '開始　出荷予定日　日
Private Const ptxE_DEN_DT_YY% = 3           '終了　出荷予定日　年
Private Const ptxE_DEN_DT_MM% = 4           '終了　出荷予定日　月
Private Const ptxE_DEN_DT_DD% = 5           '終了　出荷予定日　日
Private Const ptxS_Soko_No% = 6             '開始　倉庫№
Private Const ptxE_Soko_No% = 7             '終了　倉庫№

Private Const Text_Max% = 7                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbPRINT_KBN% = 0            '印刷区分
Private Const pcmbCyu_Kbn% = 1              '注文区分
Private Const pcmbMUKE_Code% = 2            '向け先


Private Const Print_KBN0$ = "新規　"
Private Const Print_KBN1$ = "再印刷"
Private Const Print_KBN_SIN$ = "0"
Private Const Print_KBN_SAI$ = "1"

Private KASO_NYUKA_SOKO As String * 2       '仮想　入荷倉庫番号
Private KASO_SYOHN_SOKO As String * 2       '仮想　商品化倉庫番号
Private KASO_NAI_SOKO As String * 2         '仮想　内職倉庫番号


Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）
'Dim PRT_CAN As Boolean                      '印刷途中キャンセル要求


Dim NormalFont As New StdFont               '印刷フォント
Dim Code39Font As New StdFont               '印刷フォント


Private Function Y_Syu_Get(com As Integer) As Integer

Dim sts As Integer
Dim OP  As Integer
Dim ans As Integer

    
    If com = BtOpGetGreaterEqual Then
                                        '最初のＫＥＹセット
        Call UniCode_Conv(K6_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCyu_Kbn).Text, 1))
        Call UniCode_Conv(K6_Y_SYU.HTANABAN, Text(ptxS_Soko_No).Text)
        Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
        Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    End If
    
    OP = com
    
    Do
                                        '新規の場合はＬｏｃｋ付き
        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
            OP = OP + BtSNoWait
        End If
        
        Do
            sts = BTRV(OP, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCyu_Kbn).Text, 1) Then
                                                        '事業部，完了区分，注文区分ブレーク
                        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                                Y_Syu_Get = sts
                                Exit Function
                            End If
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    
                    End If
                    If Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) > Text(ptxE_Soko_No).Text Then
                                                        '棚番オーバー
                        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                                Y_Syu_Get = sts
                                Exit Function
                            End If
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                                                        '出荷予定日
                    If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Or _
                       StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                                Y_Syu_Get = sts
                                Exit Function
                            End If
                        End If
                    Else
                                                    
                        Select Case Right(Combo(pcmbPRINT_KBN).Text, 1)
                            Case Print_KBN_SIN          '新規印刷
                                
                                If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_UN And _
                                    Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
                                                        '向け先全件指定ならＯＫ
                                    If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) = "" Then
                                        Y_Syu_Get = BtNoErr
                                        Exit Function
                                    End If
                                                        '向け先ＯＫ？
                                    If (StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)) = Right(Combo(pcmbMUKE_Code).Text, 16) Then
                                        Y_Syu_Get = BtNoErr
                                        Exit Function
                                    End If
                                End If
                            
                            Case Else                   '再印刷
                                
                                If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_UN And _
                                    Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) <> 0 Then
                                                        '向け先全件指定ならＯＫ
                                    If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) = "" Then
                                        Y_Syu_Get = BtNoErr
                                        Exit Function
                                    End If
                                                        '向け先ＯＫ？
                                    If (StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)) = Right(Combo(pcmbMUKE_Code).Text, 16) Then
                                        Y_Syu_Get = BtNoErr
                                        Exit Function
                                    End If
                                End If
                            
                        End Select
                                                        
                    End If
                                                        'データ対象外
                    If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                            Y_Syu_Get = sts
                            Exit Function
                        End If
                    End If


                    OP = BtOpGetNext
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Y_Syu_Get = BtErrEOF
                        Call Log_Out(LOG_F, "step 3")
                        Exit Function
                    End If
                Case BtErrEOF
                    Y_Syu_Get = sts
                    Exit Function
                Case Else
                    Call File_Error(sts, OP + BtSNoWait, "出荷予定ファイル")
                    Y_Syu_Get = sts
                    Exit Function
            End Select
        Loop
    Loop
End Function

                                            'エラーチェック
Private Function Err_Chk() As Integer
                                            
Dim i   As Integer
                                            
                                            
    Err_Chk = True

'注文区分
    If Combo(pcmbCyu_Kbn).Text = "" Then
        Beep
        MsgBox "注文区分を選択してください。"
        Combo(pcmbCyu_Kbn).SetFocus
        Exit Function
    End If
'出荷予定日
    For i = ptxS_DEN_DT_YY To ptxE_DEN_DT_DD
        If Len(Trim(Text(i).Text)) = 0 Then
            Select Case i
                Case ptxS_DEN_DT_YY
                    Text(i).Text = "0000"
                Case ptxS_DEN_DT_MM, ptxS_DEN_DT_DD
                    Text(i).Text = "00"
                Case ptxE_DEN_DT_YY
                    Text(i).Text = "9999"
                Case ptxE_DEN_DT_MM, ptxE_DEN_DT_DD
                    Text(i).Text = "99"
            End Select
        Else
            If IsNumeric(Trim(Text(i).Text)) Then
                Select Case i
                    Case ptxS_DEN_DT_YY, ptxE_DEN_DT_YY
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    Case Else
                        Text(i).Text = Format(CInt(Text(i).Text), "00")
                End Select
            End If
        End If
        
    Next i


    If (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) _
        > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_DEN_DT_YY).SetFocus
        Exit Function
    End If
'標準棚番(倉庫)
    If Len(Text(ptxE_Soko_No).Text) = 0 Then
        Text(ptxE_Soko_No).Text = "ZZ"
    End If

    If Text(ptxS_Soko_No).Text > Text(ptxE_Soko_No).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_Soko_No).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function
Private Function Print_Proc() As Integer

Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 2
Dim PRI_HIN_GAI     As String * 13
Dim Betu_LOCATION   As String * 8

Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long
Dim RetBuf          As String
    
    Print_Proc = True

    Call Input_Lock
    
'    PRT_CAN = False
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            '印刷中断要求
'※この処理では印刷中断は不可とする。
'        If PRT_CAN Then
'            Call Input_UnLock
'            Printer.KillDoc
'            Print_Proc = False
'            Exit Function
'        End If
                                            '出荷予定データ読み込み
        sts = Y_Syu_Get(com)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select
                                            
        If Lcnt = 99 Then
            SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
        Else
                                            '倉庫のブレーク
            If SAVE_SOKO_No <> Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) Then
                Lcnt = LMAX + 1
                SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
            End If
        End If

        If Lcnt > LMAX Then                 'ヘッダーコントロール
            If Head_Proc(Lcnt) Then
                Exit Function
            End If
            PRI_HIN_GAI = ""
        End If
                                            
        '-----------------------------------------------------  '１行目
        If StrConv(Y_SYUREC.HIN_NO, vbUnicode) <> PRI_HIN_GAI Then
            PRI_HIN_GAI = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                            '明細印刷
            Printer.Print Tab(MGN_L);
                                            '標準棚番
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);

            Printer.Print Tab(MGN_L + 10);
                                            '品番(外)
            Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);

            Printer.Print Tab(MGN_L + 24);
                                            '標準棚　在庫数
            If Len(Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode))) = 0 Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        StrConv(Y_SYUREC.HTANABAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
                       
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '別置棚検索
            If Tana_Kensaku(Betu_LOCATION) Then
                Print_Proc = True
                Exit Function
            End If
            
            SUMI_QTY = 0
            MI_QTY = 0
            
            If Len(Trim(Betu_LOCATION)) = 0 Then
            Else
                                            '別置棚　在庫数
                Printer.Print Tab(MGN_L + 35);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
                
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        Betu_LOCATION) Then
                    Exit Function
                End If
            End If
            
            Printer.Print Tab(MGN_L + 46);
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '商品化＆内職在庫数
            Printer.Print Tab(MGN_L + 55);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NAI_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            
                                            '入荷倉庫在庫
            Printer.Print Tab(MGN_L + 64);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
                        
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
        End If
        
        '2003.06.03（注文区分）
        Printer.Print Tab(MGN_L + 73);
        Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
            Case CYU_KBN_SPO
                Printer.Print " 緊";
            Case CYU_KBN_HJU
                Printer.Print " 補";
            Case Else
                Printer.Print " 　";
        End Select
        '2003.06.03
                    
                                            '伝票№
        Printer.Print Tab(MGN_L + 77);
        Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);
                                            '向け先ｺｰﾄﾞ
        Printer.Print Tab(MGN_L + 86);
        Printer.Print StrConv(Y_SYUREC.MUKE_CODE, vbUnicode);
                                            '向け先名称
        Printer.Print Tab(MGN_L + 95);
        Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
        Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Printer.Print StrConv(MTSREC.MUKE_DNAME, vbUnicode);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                Exit Function
        End Select


        Printer.Print Tab(MGN_L + 105);
        TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
        RetBuf = Format(TEMP_QTY, "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;

        Printer.Print Tab(MGN_L + 115);
                                                '印刷フォント設定（Ｃｏｄｅ３９）
        Set Printer.Font = Code39Font
                            'バーコード(*伝票ID*)
        Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                '印刷フォント設定（通常）
        Set Printer.Font = NormalFont
        
        '-----------------------------------------------------  '２行目
        Printer.Print Tab(MGN_L + 77);
        Printer.Print StrConv(Y_SYUREC.ID_NO, vbUnicode);
                                                    '向け先ｺｰﾄﾞ
        Printer.Print Tab(MGN_L + 86);
        Printer.Print StrConv(Y_SYUREC.SS_CODE, vbUnicode);

        Printer.Print
        Printer.Print
        
        Lcnt = Lcnt + 3

                                                '印刷日付設定更新
        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            
            Do
        
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Print_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定")
                        Print_Proc = SYS_ERR
                        Exit Function
                        
                End Select
        
        
            Loop
        End If
        
        
        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If



    Call Input_UnLock

    Print_Proc = False

End Function
                                    
Private Function Head_Proc(Lcnt As Integer) As Integer
Dim i As Integer
Dim sts As Integer

    Head_Proc = True

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    'Printer.Print Tab(3);                  '97.10.14
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    
'    Printer.Print Tab(MGN_L + 20); "前作業 ";
'                                        '印刷フォント設定
'    Set Printer.Font = Code39Font
'    Printer.Print "*LAST*";
'    Set Printer.Font = NormalFont
    
    Printer.Print Tab(MGN_L + 41);
    
    Printer.Print "『" + RTrim(Left(Combo(pcmbCyu_Kbn).Text, Len(Combo(pcmbCyu_Kbn).Text) - 1)) + "』出庫表";
    
    If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SAI Then
        Printer.Print Tab(MGN_L + 73);
        Printer.Print "【再印刷】";
    End If
    
    Printer.Print Tab(MGN_L + 91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L + 5);
    Printer.Print "倉庫：";
    Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2);
    Printer.Print Tab(MGN_L + 15);
    Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
            Exit Function
    End Select
    'Printer.Print                              '97.10.14
'    Printer.Print Tab(MGN_L + 90); "数量OK  ";
                                        '印刷フォント設定
'    Set Printer.Font = Code39Font
'    Printer.Print "*OK*"
'    Set Printer.Font = NormalFont
                                                '97.10.14 ここまで
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 23);
    Printer.Print "標準棚在庫";
    Printer.Print Tab(MGN_L + 35);
    Printer.Print "別置棚番";
    Printer.Print Tab(MGN_L + 47);
    Printer.Print "別置在庫";
    Printer.Print Tab(MGN_L + 56);
    Printer.Print "商品化室";
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "入荷倉庫";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "出 荷 先";
    Printer.Print Tab(MGN_L + 105);
    Printer.Print "出荷数";
    Printer.Print

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2) Then
                                                'システム倉庫の判定
                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '考えられないので読み飛ばし
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "倉庫マスタ")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
            
            
    Loop
    
    Tana_Kensaku = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030301)


    F1030301.MousePointer = vbDefault

End Sub
Private Function MTS_Set_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Edit    As String


    MTS_Set_Proc = True
    
    Call Input_Lock
    
    Combo(pcmbMUKE_Code).Clear
    
    Combo(pcmbMUKE_Code).AddItem "全出荷先" & Space(16)
    
    com = BtOpGetFirst
    
    
    Do
        
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                MTS_Set_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        Edit = StrConv(MTSREC.MUKE_DNAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        Combo(pcmbMUKE_Code).AddItem Edit
    
    
        com = BtOpGetNext
    
    Loop
    


    Call Input_UnLock
    
    Combo(pcmbMUKE_Code).ListIndex = 0

    MTS_Set_Proc = False
End Function




Private Sub Combo_Click(Index As Integer)
    Select Case Index
        Case pcmbCyu_Kbn        '注文区分
            
            If MTS_Set_Proc() Then
                Unload Me
            End If
            
            Combo(pcmbCyu_Kbn).SetFocus
    End Select

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbPRINT_KBN      '印刷区分
            Combo(pcmbCyu_Kbn).SetFocus
        Case pcmbCyu_Kbn        '注文区分
            
            If MTS_Set_Proc() Then
                Unload Me
            End If
            
            Combo(pcmbMUKE_Code).SetFocus
        Case pcmbMUKE_Code      '出荷先
            Text(ptxS_DEN_DT_YY).SetFocus
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer
    
    Select Case Index
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            yn = MsgBox("「出庫表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                For i = ZERO To Text_Max
'                    Text(i).Text = ""
'                Next i
            End If
            Combo(pcmbMUKE_Code).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
'    PRT_CAN = True
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

'
Private Sub Form_Load()

Dim c   As String * 128
Dim i   As Integer
     
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

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1030301.Caption = "標準棚番別出庫表印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '入荷仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '商品化仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '内職仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '作業管理マスタＯＰＥＮ
    If SAGYO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1030301.FontName
        .Size = 10
    End With
                                '印刷フォント設定（バーコード）
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
                                
                                '画面初期設定
    Combo(pcmbPRINT_KBN).AddItem Print_KBN0 & "   " & Print_KBN_SIN
    Combo(pcmbPRINT_KBN).AddItem Print_KBN1 & "   " & Print_KBN_SAI
    Combo(pcmbPRINT_KBN).ListIndex = 0

'    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_0 & "   " & CYU_KBN_HSP
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbCyu_Kbn).ListIndex = 0


    Text(ptxS_DEN_DT_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxS_DEN_DT_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxS_DEN_DT_DD).Text = Right(Format(Date, "yyyymmdd"), 2)

    Text(ptxE_DEN_DT_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxE_DEN_DT_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxE_DEN_DT_DD).Text = Right(Format(Date, "yyyymmdd"), 2)

    Combo(pcmbPRINT_KBN).SetFocus
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
'    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "品目マスタ")
'        End If
'    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '作業管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SAGYO_POS, SAGYOREC, Len(SAGYOREC), K0_SAGYO, Len(K0_SAGYO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "作業管理マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030301 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030301.Caption = "標準棚番別出庫表印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）"
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
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub

