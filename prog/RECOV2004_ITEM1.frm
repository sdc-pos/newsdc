VERSION 5.00
Begin VB.Form RECOV2004_ITEM1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理"
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
   Begin VB.Label Cnt 
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
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
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
      Caption         =   "データコンバート処理"
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
      Width           =   4800
   End
End
Attribute VB_Name = "RECOV2004_ITEM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim com_Zaiko       As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  品目マスタのコンバート
    MsgLab(1) = "品目マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If

        
        If Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 1) < " " Or _
            Len(Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))) = 0 Then
            
            
            Call UniCode_Conv(K0_RECOV_ITEM.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_RECOV_ITEM.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_RECOV_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
Debug.Print (StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, RECOV_ITEM_POS, RECOV_ITEMREC, Len(RECOV_ITEMREC), K0_RECOV_ITEM, Len(K0_RECOV_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(RECOV_ITEMREC.HIN_NAI, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
                Case Else
                    Call File_Error(sts, com, "元　品目マスタ")
                    Exit Function
            End Select
        
        
            Do
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                        Exit Function
                End Select
            Loop
        
            Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
            Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
            Call UniCode_Conv(K6_ZAIKO.Retu, "")
            Call UniCode_Conv(K6_ZAIKO.Ren, "")
            Call UniCode_Conv(K6_ZAIKO.Dan, "")
        
            com_Zaiko = BtOpGetGreater
        
            Do
            
                sts = BTRV(com_Zaiko, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                Select Case sts
                    Case BtNoErr
                    
                        If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                            Exit Do
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com_Zaiko, "在庫データ")
                        Exit Function
                End Select
            
                Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
            
                Do
                    sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
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
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                            Exit Function
                    End Select
                
                Loop
            
                com_Zaiko = BtOpGetNext
            Loop
        
        
        End If
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")
'---------------------------------------------  終了
Update_End:
    
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
    If RECOV_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）在庫データＯＰＥＮ
    
                    
    
    
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
  sts = BTRV(BtOpClose, RECOV_ITEM_POS, RECOV_ITEMREC, Len(RECOV_ITEMREC), K0_RECOV_ITEM, Len(K0_RECOV_ITEM), 0)
  If sts Then
      If sts <> BtErrNoOpen Then
          Call File_Error(sts, BtOpClose, "元品目マスタ")
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
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set RECOV2004_ITEM1 = Nothing

    End
End Sub

