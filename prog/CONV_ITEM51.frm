VERSION 5.00
Begin VB.Form CONV_ITEM51 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データ抽出処理（CONV_ITEM5 2010.08)"
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
   Begin VB.CommandButton Command1 
      Caption         =   "チェック＆更新"
      Height          =   435
      Index           =   1
      Left            =   7380
      TabIndex        =   15
      Top             =   1920
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   435
      Index           =   2
      Left            =   7290
      TabIndex        =   14
      Top             =   3480
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "チェック"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   11
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4725
      TabIndex        =   10
      Top             =   1620
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   9
      Top             =   1620
      Width           =   2265
   End
   Begin VB.Label Label4 
      Caption         =   "マッチング出来なかった不明なデータをDBに出力する"
      Height          =   375
      Left            =   45
      TabIndex        =   16
      Top             =   300
      Width           =   8070
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   1
      Left            =   3735
      TabIndex        =   13
      Top             =   5880
      Width           =   2445
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   0
      Left            =   1170
      TabIndex        =   12
      Top             =   5880
      Width           =   2445
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
      Caption         =   "更新件数"
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
      Caption         =   "データ抽出処理"
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
      Width           =   3360
   End
End
Attribute VB_Name = "CONV_ITEM51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc(Mode As Integer) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long


Dim c               As String * 128

Dim i               As Integer


Dim Ins_Flg         As Boolean
Dim Skip_F          As Boolean


Dim Start_Now       As String

Dim wk              As String



Dim SAVE_L_PAPER    As String
Dim MOTO_L_PAPER    As String



Dim SAVE_L_PLASTIC  As String
Dim MOTO_L_PLASTIC  As String

Dim Mark            As String * 1




    Update_Proc = True

    sts = GetIni("FILE", FUMEI_ITEM_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [FUMEI_ITEM]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)


    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0


    If FUMEI_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If



    Label3(0).Caption = Format(Now)
    Label3(1).Caption = ""


'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "品目マスタ抽出処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(Count, "#0")
    Cnt(2).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SAVE_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_SAVE_ITEM, Len(K0_SAVE_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(SAVE_ITEMREC.JGYOBU, vbUnicode) = "S" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
'        If StrConv(ITEMREC.UPD_DATETIME, vbUnicode) >= Trim(Text1(0).Text) And _
'            StrConv(ITEMREC.UPD_DATETIME, vbUnicode) <= Trim(Text1(1).Text) Then
        
        
            sel_count = sel_count + 1
            Cnt(1).Caption = Format(sel_count, "#0")
                
                
                
                
                
            Call UniCode_Conv(K0_MOTO_ITEM.JGYOBU, StrConv(SAVE_ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_MOTO_ITEM.NAIGAI, StrConv(SAVE_ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_MOTO_ITEM.HIN_GAI, StrConv(SAVE_ITEMREC.HIN_GAI, vbUnicode))
                
            sts = BTRV(BtOpGetEqual, MOTO_ITEM_POS, MOTO_ITEMREC, Len(MOTO_ITEMREC), K0_MOTO_ITEM, Len(K0_MOTO_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Skip_F = False
                
                    If Trim(StrConv(SAVE_ITEMREC.L_JGYOBU_CODE, vbUnicode)) = "" Then
                        Skip_F = True
                    End If
                    If Trim(StrConv(SAVE_ITEMREC.L_KAISHA_CODE, vbUnicode)) = "" Then
                        Skip_F = True
                    End If
                
                
                    If Not IsNumeric(StrConv(SAVE_ITEMREC.L_URIKIN1, vbUnicode)) Then
                        Skip_F = True
                    End If
                
                    If Not IsNumeric(StrConv(SAVE_ITEMREC.L_URIKIN2, vbUnicode)) Then
                        Skip_F = True
                    End If
                
                
                    If Not IsNumeric(StrConv(SAVE_ITEMREC.L_URIKIN3, vbUnicode)) Then
                        Skip_F = True
                    End If
                
                
                
                
                
                    If Skip_F Then
                        Mark = "*"
                    Else
                        Mark = ""
                    End If
                

                    If StrConv(SAVE_ITEMREC.L_PAPER, vbUnicode) <> "1" Then
                        SAVE_L_PAPER = "0"
                    End If
                    
                    If StrConv(MOTO_ITEMREC.L_PAPER, vbUnicode) <> "1" Then
                        MOTO_L_PAPER = "0"
                    End If
                    
                    
                    
                    If StrConv(SAVE_ITEMREC.L_PLASTIC, vbUnicode) <> "1" Then
                        SAVE_L_PLASTIC = "0"
                    End If
                    
                    If StrConv(MOTO_ITEMREC.L_PLASTIC, vbUnicode) <> "1" Then
                        MOTO_L_PLASTIC = "0"
                    End If
                    
                    
                    
                    If SAVE_L_PAPER <> MOTO_L_PAPER Or _
                        SAVE_L_PLASTIC <> MOTO_L_PLASTIC Then
                        
                                    
                        
                        
                        If StrConv(SAVE_ITEMREC.L_PAPER, vbUnicode) = "0" And StrConv(SAVE_ITEMREC.L_PLASTIC, vbUnicode) = "0" Then
                        
                        
                        
                        
                        
                        
                        
                        
                            sts = BTRV(BtOpInsert, FUMEI_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_FUMEI_ITEM, Len(K0_FUMEI_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrDuplicates
                                Case Else
                                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                                    Exit Function
                            End Select
                        
                        
                        
                            Call LOG_OUT(Start_Now & "item.txt", "," & "UPD," & StrConv(SAVE_ITEMREC.JGYOBU, vbUnicode) & "," & _
                                        StrConv(SAVE_ITEMREC.NAIGAI, vbUnicode) & "," & _
                                        StrConv(SAVE_ITEMREC.HIN_GAI, vbUnicode) & "," & _
                                        "旧紙=" & StrConv(MOTO_ITEMREC.L_PAPER, vbUnicode) & "," & _
                                        "旧プラ=" & StrConv(MOTO_ITEMREC.L_PLASTIC, vbUnicode) & "," & _
                                        "更新担当=" & StrConv(SAVE_ITEMREC.UPD_TANTO, vbUnicode) & "," & _
                                        "更新日時=" & StrConv(SAVE_ITEMREC.UPD_DATETIME, vbUnicode) & "," & Mark)
                            
                            
                            
                            If Mode = 1 Then
                            
                                Call UniCode_Conv(SAVE_ITEMREC.L_PAPER, StrConv(MOTO_ITEMREC.L_PAPER, vbUnicode))
                                Call UniCode_Conv(SAVE_ITEMREC.L_PLASTIC, StrConv(MOTO_ITEMREC.L_PLASTIC, vbUnicode))
                        
                                sts = BTRV(BtOpUpdate, SAVE_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_SAVE_ITEM, Len(K0_SAVE_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Exit Do
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "品目マスタ")
                                        Exit Function
                                End Select
                                        
                            End If
                        
                        
                        End If
                        
                        
                        upd_count = upd_count + 1
                        Cnt(2).Caption = Format(upd_count, "#0")
                        
                        
                        
                    End If
                
                
                
                
                
                
                
                
                
                
                
                Case BtErrKeyNotFound
                
                
                
                
                    sts = BTRV(BtOpInsert, FUMEI_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_FUMEI_ITEM, Len(K0_FUMEI_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrDuplicates
                        Case Else
                            Call File_Error(sts, BtOpInsert, "品目マスタ")
                            Exit Function
                    End Select
                
                
                
                    Call LOG_OUT(Start_Now & "item.txt", "," & "INS," & StrConv(SAVE_ITEMREC.JGYOBU, vbUnicode) & "," & _
                                StrConv(SAVE_ITEMREC.NAIGAI, vbUnicode) & "," & _
                                StrConv(SAVE_ITEMREC.HIN_GAI, vbUnicode) & "," & _
                                "紙=" & StrConv(SAVE_ITEMREC.L_PAPER, vbUnicode) & "," & _
                                "プラ=" & StrConv(SAVE_ITEMREC.L_PLASTIC, vbUnicode) & "," & _
                                "更新担当=" & StrConv(SAVE_ITEMREC.UPD_TANTO, vbUnicode) & "," & _
                                "更新日時=" & StrConv(SAVE_ITEMREC.UPD_DATETIME, vbUnicode))
                
                
                
                
                
                
                    upd_count = upd_count + 1
                    Cnt(2).Caption = Format(upd_count, "#0")
                
                
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            '抽出条件
            
            
                
        
'        End If
        
        
        com = BtOpGetNext
    
    Loop


    sts = BTRV(BtOpClose, FUMEI_ITEM_POS, FUMEI_ITEMREC, Len(FUMEI_ITEMREC), K0_FUMEI_ITEM, Len(K0_FUMEI_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If







    Cnt(0).Caption = Format(Count, "#0")

    Label3(1).Caption = Format(Now)
    Me.MousePointer = vbDefault

'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function




Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
Dim sts As Integer
    
Dim FullPath    As String
Dim c           As String * 128


    
    
    
    
    
    Select Case Index
        Case 0
            ans = MsgBox("「チェック」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                
                If Update_Proc(0) Then
                    Unload Me
                End If
            
            
            
            
                MsgBox "終了しました"
            
            End If


        Case 1

            ans = MsgBox("「チェック＆更新」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                
                If Update_Proc(1) Then
                    Unload Me
                End If
            
            
            
            
                MsgBox "終了しました"
            
            End If




        Case 2
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
    Text1(0).Text = "20100716164000"
    Text1(1).Text = "20101231235959"
                                
                                

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
    
    If SAVE_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
    If MOTO_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
                    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, SAVE_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_SAVE_ITEM, Len(K0_SAVE_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, SAVE_ITEM_POS, SAVE_ITEMREC, Len(SAVE_ITEMREC), K0_SAVE_ITEM, Len(K0_SAVE_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_ITEM41 = Nothing

    End
End Sub

