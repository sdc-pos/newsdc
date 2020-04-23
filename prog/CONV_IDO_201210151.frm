VERSION 5.00
Begin VB.Form CONV_IDO_201210151 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理（CONV_ZAIKO_20121009 2012.10.10 08:45)"
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
   Begin VB.CheckBox Check1 
      Caption         =   "ログ出力"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   10
      Top             =   1680
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   9
      Top             =   1080
      Width           =   1860
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   8
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "書換件数"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "対象件数"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "読込み件数"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   2
      Top             =   4800
      Width           =   1410
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
      Caption         =   "データリカバリー処理"
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
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4800
   End
End
Attribute VB_Name = "CONV_IDO_201210151"
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
Dim Count           As Long
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long

Dim wkKISHU1        As String * 25
Dim wkKISHU2        As String * 52
Dim wkKISHU3        As String * 150
Dim wkKISHU_BIKOU   As String * 450

Dim c               As String * 128

Dim i               As Integer

Dim CHG_FLG         As Boolean

Dim Start_Now       As String


Dim wkSoko_No       As String * 2   '倉庫№
Dim wkRetu          As String * 2   '棚番　列
Dim wkRen           As String * 2   '棚番　連
Dim wkDan           As String * 2   '棚番　段
Dim wkJGYOBU        As String * 1   '事業部区分
Dim wkNAIGAI        As String * 1   '国内外
Dim wkHIN_GAI       As String * 20  '品番（外部）
Dim wkGOODS_ON      As String * 1   '商品化／未商品化
Dim wkNYUKA_DT      As String * 8   '入荷日付

Dim wkNYUKO_DT      As String * 8   '入庫日付


Dim wkHIN_NAI       As String * 20

Dim wkYUKO_Z_QTY    As String * 8   '有効在庫数
Dim YUKO_Z_QTY      As Long
    Update_Proc = True


'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "在庫移動歴データリカバリー処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(Count, "#0")
    Cnt(2).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
    For i = 0 To 19
    
        Mid(wkHIN_NAI, i + 1, 1) = vbNullChar
    Next i
                                        
                                        
    Call UniCode_Conv(K1_IDO.JGYOBU, "B")
    Call UniCode_Conv(K1_IDO.NAIGAI, "1")
    Call UniCode_Conv(K1_IDO.HIN_GAI, "AD-KT37K3F-C")
    Call UniCode_Conv(K1_IDO.JITU_DT, "20110101")
    Call UniCode_Conv(K1_IDO.JITU_TM, "")
    
                                        
    com = BtOpGetGreater
    Do
        
        DoEvents
        
        
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            
                If Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) > "AD-KT37K3F-C" Then
                    Exit Do
                End If
            
                            
            
            
                If StrConv(IDOREC.JITU_DT, vbUnicode) > "20110331" Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
'        If Trim(StrConv(IDOREC.NYUKA_DT, vbUnicode, vbUnicode)) <> "" Then
        upd_count = upd_count + 1
        Cnt(1).Caption = Format(upd_count, "#0")
        DoEvents
            
            
                Call UniCode_Conv(IDOREC.HIN_NAI, wkHIN_GAI)
            
                Call UniCode_Conv(IDOREC.HIN_NAI, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.TANTO_CODE, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.TANTO_NAME, wkHIN_NAI)
                            
                Call UniCode_Conv(IDOREC.MUKE_CODE, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.MUKE_DNAME, wkHIN_NAI)
                            
                Call UniCode_Conv(IDOREC.MEMO, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.DEN_NO, wkHIN_NAI)
                            
                Call UniCode_Conv(IDOREC.ID_NO, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.WEL_ID, wkHIN_NAI)
                            
                Call UniCode_Conv(IDOREC.SHIIRE_CODE, wkHIN_NAI)
                Call UniCode_Conv(IDOREC.SHIIRE_TANKA, wkHIN_NAI)
                            
                            
            sts = BTRV(BtOpUpdate, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpDelete, "在庫データ")
                    Exit Function
            End Select
'        End If
        
                
                                    
                
        com = BtOpGetNext
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
    Select Case Index
        Case 0
            ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                MsgBox "終了しました"
                Unload Me
            
            End If


        Case 1
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
                                
                                

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
    
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    Check1.Value = vbChecked
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_IDO_201210151 = Nothing

    End
End Sub

