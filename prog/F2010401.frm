VERSION 5.00
Begin VB.Form F2010401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出荷明細データ出力"
   ClientHeight    =   7080
   ClientLeft      =   2325
   ClientTop       =   2910
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
   ScaleHeight     =   7080
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   11
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   31
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   10
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   30
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   9
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   29
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   8
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   28
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   27
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   26
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "終　了"
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
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
      Index           =   8
      Left            =   7800
      TabIndex        =   14
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
      TabIndex        =   7
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   11
      Left            =   4680
      TabIndex        =   36
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   10
      Left            =   5160
      TabIndex        =   35
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "～"
      Height          =   252
      Index           =   9
      Left            =   5760
      TabIndex        =   34
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   8
      Left            =   6720
      TabIndex        =   33
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   7
      Left            =   7200
      TabIndex        =   32
      Top             =   2760
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "対象伝票日付"
      Height          =   252
      Index           =   6
      Left            =   2520
      TabIndex        =   25
      Top             =   2760
      Width           =   1452
      WordWrap        =   -1  'True
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
      TabIndex        =   24
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   5
      Left            =   7200
      TabIndex        =   23
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   4
      Left            =   6720
      TabIndex        =   22
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "～"
      Height          =   252
      Index           =   3
      Left            =   5760
      TabIndex        =   21
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   2
      Left            =   5160
      TabIndex        =   20
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "/"
      Height          =   252
      Index           =   1
      Left            =   4680
      TabIndex        =   19
      Top             =   2040
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "集計年月日"
      Height          =   252
      Index           =   0
      Left            =   2760
      TabIndex        =   18
      Top             =   2040
      Width           =   1212
      WordWrap        =   -1  'True
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
Attribute VB_Name = "F2010401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_JITU_DT_YY% = 0          '開始集計年月日 年
Private Const ptxS_JITU_DT_MM% = 1          '開始集計年月日 月
Private Const ptxS_JITU_DT_DD% = 2          '開始集計年月日 日
Private Const ptxE_JITU_DT_YY% = 3          '終了集計年月日 年
Private Const ptxE_JITU_DT_MM% = 4          '終了集計年月日 月
Private Const ptxE_JITU_DT_DD% = 5          '終了集計年月日 日
Private Const ptxS_DEN_DT_YY% = 6           '開始伝票年月日 年
Private Const ptxS_DEN_DT_MM% = 7           '開始伝票年月日 月
Private Const ptxS_DEN_DT_DD% = 8           '開始伝票年月日 日
Private Const ptxE_DEN_DT_YY% = 9           '終了伝票年月日 年
Private Const ptxE_DEN_DT_MM% = 10          '終了伝票年月日 月
Private Const ptxE_DEN_DT_DD% = 11          '終了伝票年月日 日

Private Const Text_Max% = 11                '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim MEIJITU_DATA  As String                 '入出荷明細データフルパス
Private Function Main_Proc() As Integer
                                 
Dim c               As String * 128

Dim sts             As Integer
Dim com             As Integer
Dim Upd_Com         As Integer
Dim ans             As Integer
Dim Ret             As Integer

Dim FileNo          As Integer
Dim fileName        As String
    
    Main_Proc = True
                                 
    Call Input_Lock
                                 '入出荷明細ファイル削除
    If GetIni("FILE", MEIJ_ID, "SYS", c) Then
        Beep
        MsgBox "入出荷明細ファイル情報の獲得に失敗しました。処理を中止して下さい。"
        Call Log_Out(LOG_F, "[SYS.INI] [FILE] [MEIJ] READ ERROR")
        Exit Function
    End If
    
    On Error Resume Next
    Kill RTrim(c)
                                
                                '入出荷明細ファイルＯＰＥＮ
    If MEIJ_Open(BtOpenNomal) Then
        Exit Function
    End If
                                            '移動歴より集計
    Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxS_JITU_DT_YY).Text & Text(ptxS_JITU_DT_MM).Text & Text(ptxS_JITU_DT_DD).Text)
    Call UniCode_Conv(K0_IDO.JITU_TM, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
                If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxE_JITU_DT_YY).Text & Text(ptxE_JITU_DT_MM).Text & Text(ptxE_JITU_DT_DD).Text) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                Exit Function
        End Select
        
        Select Case StrConv(IDOREC.RIRK_ID, vbUnicode)
            Case YOIN_TU_NYUKA
                                '入荷
                                
                                '入出荷明細ファイル読込み
                Call UniCode_Conv(K0_MEIJ.IO_KBN, "1")
                Call UniCode_Conv(K0_MEIJ.DEN_DT, StrConv(IDOREC.NYUKA_DT, vbUnicode))
                Call UniCode_Conv(K0_MEIJ.CYU_KBN, "")
                Call UniCode_Conv(K0_MEIJ.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_MEIJ.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
            
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_Com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_Com = BtOpInsert
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE 'ここではない！
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<MEIJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入出荷明細ファイル")
                            Exit Function
                    End Select
                Loop
        
                If Upd_Com = BtOpInsert Then
                    Call UniCode_Conv(MEIJREC.IO_KBN, "1")
                    Call UniCode_Conv(MEIJREC.DEN_DT, StrConv(IDOREC.NYUKA_DT, vbUnicode))
                    Call UniCode_Conv(MEIJREC.CYU_KBN, "")
                    Call UniCode_Conv(MEIJREC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(MEIJREC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(MEIJREC.JITU_QTY, Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                Else
                    Call UniCode_Conv(MEIJREC.JITU_QTY, Format(CLng(StrConv(MEIJREC.JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                End If
        
            
                Do
                    sts = BTRV(Upd_Com, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<MEIJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, Upd_Com, "入出荷明細ファイル")
                            Exit Function
                    End Select
            
                Loop
            
            
            
            Case (ACT_SYUKA_KEI & CYU_KBN_HSP), _
                    (ACT_SYUKA_KEI & CYU_KBN_TUK), _
                    (ACT_SYUKA_KEI & CYU_KBN_SPO), _
                    (ACT_SYUKA_KEI & CYU_KBN_HJU), _
                    (ACT_SYUKA_HYO & CYU_KBN_HSP), _
                    (ACT_SYUKA_HYO & CYU_KBN_TUK), _
                    (ACT_SYUKA_HYO & CYU_KBN_SPO), _
                    (ACT_SYUKA_HYO & CYU_KBN_HJU), _
                    (ACT_SYUKA_GAI & CYU_KBN_KIN)
                                '出荷（除く貿易）
                If StrConv(IDOREC.DEN_DT, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Or _
                    StrConv(IDOREC.DEN_DT, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                    Else
                                '入出荷明細ファイル読込み
                        Call UniCode_Conv(K0_MEIJ.IO_KBN, "0")
                        Call UniCode_Conv(K0_MEIJ.DEN_DT, StrConv(IDOREC.DEN_DT, vbUnicode))
                        Call UniCode_Conv(K0_MEIJ.CYU_KBN, Right(StrConv(IDOREC.RIRK_ID, vbUnicode), 1))
                        Call UniCode_Conv(K0_MEIJ.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_MEIJ.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                        Do
                            sts = BTRV(BtOpGetEqual + BtSNoWait, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
                            Select Case sts
                                Case BtNoErr
                                    Upd_Com = BtOpUpdate
                                    Exit Do
                                Case BtErrKeyNotFound
                                    Upd_Com = BtOpInsert
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE 'ここではない！
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<MEIJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "入出荷明細ファイル")
                                    Exit Function
                            End Select
                        Loop
        
                        If Upd_Com = BtOpInsert Then
                            Call UniCode_Conv(MEIJREC.IO_KBN, "0")
                            Call UniCode_Conv(MEIJREC.DEN_DT, StrConv(IDOREC.DEN_DT, vbUnicode))
                            Call UniCode_Conv(MEIJREC.CYU_KBN, Right(StrConv(IDOREC.RIRK_ID, vbUnicode), 1))
                            Call UniCode_Conv(MEIJREC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                            Call UniCode_Conv(MEIJREC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                            Call UniCode_Conv(MEIJREC.JITU_QTY, Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                        Else
                            Call UniCode_Conv(MEIJREC.JITU_QTY, Format(CLng(StrConv(MEIJREC.JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "00000000"))
                        End If
                
                        Do
                            sts = BTRV(Upd_Com, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<MEIJITU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, Upd_Com, "入出荷明細ファイル")
                                    Exit Function
                            End Select
            
                        Loop
                End If
                
        End Select
        
        
        
        com = BtOpGetNext
    
    Loop
                                            
                                            

    FileNo = FreeFile
    fileName = MEIJITU_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo
                                                
    Write #FileNo, , "伝票日付", "注文区分", , "品番（外部）", "実績数"

    com = BtOpGetFirst

    Do
        DoEvents
        sts = BTRV(com, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入出荷明細ファイル")
                Exit Function
        End Select
    
        If StrConv(MEIJREC.IO_KBN, vbUnicode) = "0" Then
            Write #FileNo, "出荷",
            
        Else
            Write #FileNo, "入荷",
        End If
    
        Write #FileNo, Left(StrConv(MEIJREC.DEN_DT, vbUnicode), 4) & "/" & Mid(StrConv(MEIJREC.DEN_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(MEIJREC.DEN_DT, vbUnicode), 2),
    
        If StrConv(MEIJREC.IO_KBN, vbUnicode) = "0" Then
            Select Case StrConv(MEIJREC.CYU_KBN, vbUnicode)
                Case CYU_KBN_HSP
                    Write #FileNo, "ス／補",
                Case CYU_KBN_TUK
                    Write #FileNo, "月切り",
                Case CYU_KBN_SPO
                    Write #FileNo, "スポット",
                Case CYU_KBN_HJU
                    Write #FileNo, "補充",
            End Select
        Else
            Write #FileNo, ,
        End If
    
    
        If StrConv(MEIJREC.NAIGAI, vbUnicode) = NAIGAI_NAI Then
            Write #FileNo, ,
        Else
            Write #FileNo, "国外",
        End If
    
        Write #FileNo, StrConv(MEIJREC.HIN_GAI, vbUnicode),
        Write #FileNo, Format(CLng(StrConv(MEIJREC.JITU_QTY, vbUnicode)), "#0")
    
        com = BtOpGetNext
    Loop
    
    Close #FileNo
    Call Input_UnLock
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    Main_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Main_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Main_Proc = True
    End If
End Function

Private Function Err_Chk()
    
Dim i As Integer
    
    Err_Chk = True


    For i = ptxS_JITU_DT_YY To ptxE_DEN_DT_DD
        If Len(Text(i).Text) = 0 Then
            Select Case i
                Case ptxS_JITU_DT_YY, ptxS_DEN_DT_YY
                    Text(i).Text = "0000"
                Case ptxE_JITU_DT_YY, ptxE_DEN_DT_YY
                    Text(i).Text = "9999"
                Case ptxS_JITU_DT_MM, ptxS_JITU_DT_DD, ptxS_DEN_DT_MM, ptxS_DEN_DT_DD
                    Text(i).Text = "00"
                Case ptxE_JITU_DT_MM, ptxE_JITU_DT_DD, ptxE_DEN_DT_MM, ptxE_DEN_DT_DD
                    Text(i).Text = "99"
            End Select
        Else
            If IsNumeric(Text(i).Text) Then
                Select Case i
                    Case ptxS_JITU_DT_YY, ptxE_JITU_DT_YY, ptxS_DEN_DT_YY, ptxE_DEN_DT_YY
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    Case Else
                        Text(i).Text = Format(CInt(Text(i).Text), "00")
                End Select
            End If
        End If
    Next i
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F2010401.MousePointer = vbHourglass

    Call Ctrl_Lock(F2010401)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F2010401)


    F2010401.MousePointer = vbDefault

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「入出荷明細データ出力」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Main_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(ptxS_JITU_DT_YY).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub

Private Sub Form_Activate()
    Text(ptxS_JITU_DT_YY).SetFocus

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

Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
        

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
                                '入出荷明細ファイル名取り込み
    If GetIni("FILE", "MEIJITU_DATA", "SYS", c) Then
        Beep
        MsgBox "入出荷明細ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    MEIJITU_DATA = Trim(c)
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
            F2010401.Caption = "入出荷明細データ出力（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If


    Text(ptxS_DEN_DT_YY).Text = Left(Format(DateAdd("d", -1, Now), "yyyymmdd"), 4)
    Text(ptxS_DEN_DT_MM).Text = Mid(Format(DateAdd("d", -1, Now), "yyyymmdd"), 5, 2)
    Text(ptxS_DEN_DT_DD).Text = Right(Format(DateAdd("d", -1, Now), "yyyymmdd"), 2)
    Text(ptxE_DEN_DT_YY).Text = Left(Format(DateAdd("d", -1, Now), "yyyymmdd"), 4)
    Text(ptxE_DEN_DT_MM).Text = Mid(Format(DateAdd("d", -1, Now), "yyyymmdd"), 5, 2)
    Text(ptxE_DEN_DT_DD).Text = Right(Format(DateAdd("d", -1, Now), "yyyymmdd"), 2)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '入出荷明細ＣＬＯＳＥ
    sts = BTRV(BtOpClose, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入出荷明細データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, MEIJ_POS, MEIJREC, Len(MEIJREC), K0_MEIJ, Len(K0_MEIJ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F2010401 = Nothing

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
    F2010401.Caption = "入出荷明細データ出力（" + RTrim(JGYOBU_T(Index).NAME) + "）"
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


