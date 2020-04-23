VERSION 5.00
Begin VB.Form F1010251 
   BackColor       =   &H00C0C0C0&
   Caption         =   "棚ファイル　使用状況セットアップ（テスト用）　F101025 2010.12.14 16:00"
   ClientHeight    =   5388
   ClientLeft      =   2328
   ClientTop       =   2628
   ClientWidth     =   9096
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
   ScaleHeight     =   5388
   ScaleWidth      =   9096
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   492
      Index           =   1
      Left            =   7200
      TabIndex        =   23
      Top             =   1800
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   492
      Index           =   0
      Left            =   7200
      TabIndex        =   22
      Top             =   1080
      Width           =   1452
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   21
      Top             =   3360
      Width           =   612
   End
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   19
      Top             =   3360
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   17
      Top             =   3360
      Width           =   1092
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "-"
      Height          =   252
      Index           =   6
      Left            =   3480
      TabIndex        =   25
      Top             =   4560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   8
      Left            =   3120
      TabIndex        =   24
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "使用状況（％）"
      Height          =   252
      Index           =   1
      Left            =   5280
      TabIndex        =   20
      Top             =   3480
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "～"
      Height          =   252
      Left            =   3480
      TabIndex        =   18
      Top             =   3480
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "棚番範囲"
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   16
      Top             =   3480
      Width           =   972
   End
   Begin VB.Label Lab_dsp 
      Height          =   252
      Index           =   7
      Left            =   5520
      TabIndex        =   15
      Top             =   4560
      Width           =   972
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "現在"
      Height          =   252
      Index           =   5
      Left            =   2400
      TabIndex        =   14
      Top             =   4560
      Width           =   492
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "最大"
      Height          =   252
      Index           =   4
      Left            =   2400
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   6
      Left            =   4920
      TabIndex        =   12
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "-"
      Height          =   252
      Index           =   3
      Left            =   4680
      TabIndex        =   11
      Top             =   4560
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   5
      Left            =   4320
      TabIndex        =   10
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "-"
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   4560
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   4
      Left            =   3720
      TabIndex        =   8
      Top             =   4560
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "-"
      Height          =   252
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '中央揃え
      Caption         =   "-"
      Height          =   252
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '中央揃え
      Height          =   252
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   2532
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "棚データ更新中！"
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
      Width           =   3840
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "環境ファイルセットアップ"
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
      Width           =   5760
   End
End
Attribute VB_Name = "F1010251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                                            '棚マスタの追加／訂正
Private Function Update_Proc() As Integer
Dim sts As Integer
Dim Upd_com As Integer
Dim com As Integer
Dim Retu, Ren, Dan As Integer
Dim SOKO   As Integer
Dim ans As Integer
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
    Me.MousePointer = vbHourglass
    
                                            
                                            
    Call UniCode_Conv(K0_SOKO.Soko_No, Mid(Text1(0).Text, 1, 2))
    com = BtOpGetGreaterEqual
    Do
        
        DoEvents
        
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.Soko_No, vbUnicode) > Mid(Text1(1).Text, 1, 2) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "倉庫マスタ")
                Update_Proc = True
                Exit Function
        End Select
        
    
    
                                        '列のループ
        For Retu = Val(Mid(Text1(0).Text, 3, 2)) To Val(Mid(Text1(1).Text, 3, 2))
                                        '連のループ
            For Ren = Val(Mid(Text1(0).Text, 5, 2)) To Val(Mid(Text1(1).Text, 5, 2))
                                        '段のループ
                For Dan = Val(Mid(Text1(0).Text, 7, 2)) To Val(Mid(Text1(1).Text, 7, 2))
    
    
                    Lab_dsp(8) = StrConv(SOKOREC.Soko_No, vbUnicode)
    
                    Lab_dsp(4) = Format(Retu, "00")
                    Lab_dsp(5) = Format(Ren, "00")
                    Lab_dsp(6) = Format(Dan, "00")
                    DoEvents            'ちょっと他プロセスの様子を見る
    
    
                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                            If Not IsNumeric(Text1(2).Text) Then
                            
                                If Trim(Text1(2).Text) = "." Then
                                    Call UniCode_Conv(TANAREC.Tana_Use, ".  ")
                                Else
                                    Call UniCode_Conv(TANAREC.Tana_Use, Format(100, "000"))
                                End If
                            Else
                                Call UniCode_Conv(TANAREC.Tana_Use, Format(Val(Text1(2).Text), "000"))
                            
                            End If
            
            
            
            
            
            
                            sts = BTRV(BtOpUpdate, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "棚マスタ")
                                    Update_Proc = True
                                    Exit Function
                            End Select
                        
                        
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "棚マスタ")
                            Update_Proc = True
                            Exit Function
                    End Select
    
    
    
                Next Dan
    
            Next Ren
    
        Next Retu
        
        com = BtOpGetNext
    
    Loop
                                            
    MsgBox "終了しました！"
                                            
    Me.MousePointer = vbDefault
                                            
                                            
                                            
End Function


Private Sub Command1_Click(Index As Integer)

    Select Case Index
    
        Case 0
            
            If Trim(Text1(0).Text) = "" Then
                Text1(0).Text = "  000000"
            End If
            If Trim(Text1(1).Text) = "" Then
                Text1(1).Text = "zz999999"
            End If
            
            
            If Text1(0).Text > Text1(1).Text Then
                MsgBox "入力エラー"
                Text1(0).SetFocus
                Exit Sub
            End If
            
            If Update_Proc() Then
                Unload Me
            End If
        Case 1
            Unload Me
    End Select


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
                                

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
    
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
    
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010251 = Nothing

    End
End Sub

