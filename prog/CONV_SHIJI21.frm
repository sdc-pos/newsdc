VERSION 5.00
Begin VB.Form CONV_SHIJI21 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データ抽出処理（CONV_ITEM2 2010.08)"
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
   Begin VB.CheckBox Check2 
      Caption         =   "プラ"
      Height          =   240
      Left            =   3420
      TabIndex        =   18
      Top             =   3420
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "紙"
      Height          =   240
      Left            =   2250
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CSV"
      Height          =   435
      Index           =   2
      Left            =   7245
      TabIndex        =   16
      Top             =   2640
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "追加"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   15
      Top             =   1860
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   435
      Index           =   3
      Left            =   7290
      TabIndex        =   12
      Top             =   3600
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新規"
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
   Begin VB.Label Label3 
      Height          =   315
      Index           =   1
      Left            =   3735
      TabIndex        =   14
      Top             =   5880
      Width           =   2445
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   0
      Left            =   1170
      TabIndex        =   13
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
      Caption         =   "抽出件数"
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
Attribute VB_Name = "CONV_SHIJI21"
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


Dim c               As String * 128

Dim i               As Integer


Dim Ins_Flg         As Boolean

Dim Start_Now       As String

Dim wk              As String








    Update_Proc = True


    Label3(0).Caption = Format(Now)
    Label3(1).Caption = ""


'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "指図票データ抽出処理中！！"
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
        
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "指図票マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
        If StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode) >= Trim(Text1(0).Text) And _
            StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode) <= Trim(Text1(1).Text) Then
        
        
            sel_count = sel_count + 1
            Cnt(1).Caption = Format(sel_count, "#0")
            
                
                
            '抽出条件
            
            Ins_Flg = True
            
'            If StrConv(ITEMREC.L_PAPER, vbUnicode) <> "0" Or StrConv(ITEMREC.L_PLASTIC, vbUnicode) <> "0" Then
'                Ins_Flg = False
'            End If
                
                
                
                
            
            
            
            
            
            
'            If Check2.Value = vbChecked Then
'                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) <> "1" Then
'                    Ins_Flg = False
'                End If
'
'            Else
'
'                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) <> "0" Then
'                    Ins_Flg = False
'                End If
'
'
'            End If
                
                
                
                
                
            If Ins_Flg Then
            
            
            
            
                Do
                    sts = BTRV(BtOpInsert, SAVE_P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_SAVE_P_SSHIJI_O, Len(K0_SAVE_P_SSHIJI_O), 0)
'                    sts = BtNoErr
                    Select Case sts
                        Case BtNoErr, BtErrDuplicates
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpInsert + BtSNoWait, "指図票ﾃﾞｰﾀ")
                            Exit Function
                    End Select
                Loop
            
            
            
'                Call LOG_OUT(Start_Now & "item.txt", StrConv(ITEMREC.JGYOBU, vbUnicode) & "," & StrConv(ITEMREC.NAIGAI, vbUnicode) & "," & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "," & wk)
                upd_count = upd_count + 1
                Cnt(2).Caption = Format(upd_count, "#0")
            
            
            
            End If
        
        
        
        End If
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")

    Label3(1).Caption = Format(Now)
    Me.MousePointer = vbDefault

'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function
Private Function Output_Proc() As Integer

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

Dim Start_Now       As String

Dim wk              As String

    
Dim Ret             As String
Dim fileName        As String
Dim FileNo          As Long
    
    







    Output_Proc = True


    Label3(0).Caption = Format(Now)
    Label3(1).Caption = ""


'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "指図票ﾃﾞｰﾀ抽出処理中！！"
    Me.MousePointer = vbHourglass
    
    
    
    
    
    If GetIni("FILE", "P_SSHIJI_O_CSV", "SYS", c) Then
        Beep
        MsgBox "指図票データ出力用ファイル[P_SSHIJI_O_CSV]の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    
    FileNo = FreeFile
    fileName = Trim(c)
    
    
    Ret = InStrRev(Trim(fileName), ".") - 1
    
    fileName = Left(Trim(fileName), Ret) & Right(Trim(fileName), Len(Trim(fileName)) - Ret)


    Open (fileName) For Output As FileNo
    
    
    Write #FileNo, "事業部", "内外", "品番（外部）", "発行日", "完了日"

    
    
    
    
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
        
        
        sts = BTRV(com, SAVE_P_SSHIJI_O_POS, SAVE_P_SSHIJI_O_REC, Len(SAVE_P_SSHIJI_O_REC), K0_SAVE_P_SSHIJI_O, Len(K0_SAVE_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
        
        
        sel_count = sel_count + 1
        Cnt(1).Caption = Format(sel_count, "#0")
            
                
                
            
            
            
        upd_count = upd_count + 1
        Cnt(2).Caption = Format(upd_count, "#0")
            
            
        Write #FileNo, StrConv(SAVE_P_SSHIJI_O_REC.JGYOBU, vbUnicode), _
                        StrConv(SAVE_P_SSHIJI_O_REC.NAIGAI, vbUnicode), _
                        Trim(StrConv(SAVE_P_SSHIJI_O_REC.HIN_GAI, vbUnicode)), _
                        Trim(StrConv(SAVE_P_SSHIJI_O_REC.HAKKO_DT, vbUnicode)), _
                        Trim(StrConv(SAVE_P_SSHIJI_O_REC.KAN_DT, vbUnicode))
            
            
        
        
        com = BtOpGetNext
    
    Loop

    Close #FileNo


    Cnt(0).Caption = Format(Count, "#0")

    Label3(1).Caption = Format(Now)
    Me.MousePointer = vbDefault

'---------------------------------------------  終了
Update_End:
    
    Output_Proc = False

End Function




Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
Dim sts As Integer
    
Dim FullPath    As String
Dim c           As String * 128


                                            '品目マスタフルパス取込み
    sts = GetIni("FILE", SAVE_P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [SAVE_P_SSHIJI_O_ITEM]読み込みエラー ")
        
        MsgBox "SYS.INI [SAVE_P_SSHIJI_O_ITEM]読み込みエラー "
        
        Exit Sub
    End If

    FullPath = RTrim(c)
    
    
    
    
    
    Select Case Index
        Case 0
            ans = MsgBox("「新規」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                On Error Resume Next
                Kill (FullPath)
                On Error GoTo 0
                
                
                
                
                
                If SAVE_P_SSHIJI_O_Open(BtOpenNomal) Then
                    Unload Me
                End If
                
                
                
                If Update_Proc() Then
                    Unload Me
                End If
            
            
                sts = BTRV(BtOpClose, SAVE_P_SSHIJI_O_POS, SAVE_P_SSHIJI_O_REC, Len(SAVE_P_SSHIJI_O_REC), K0_SAVE_P_SSHIJI_O, Len(K0_SAVE_P_SSHIJI_O), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
                    End If
                End If
            
            
                MsgBox "終了しました"
            
            End If



        Case 1
            
            
            
            
            ans = MsgBox("「追加」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                If SAVE_P_SSHIJI_O_Open(BtOpenNomal) Then
                    Unload Me
                End If
                
                
                
                If Update_Proc() Then
                    Unload Me
                End If
            
            
                sts = BTRV(BtOpClose, SAVE_P_SSHIJI_O_POS, SAVE_P_SSHIJI_O_REC, Len(SAVE_P_SSHIJI_O_REC), K0_SAVE_P_SSHIJI_O, Len(K0_SAVE_P_SSHIJI_O), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
                    End If
                End If
            
            
            
                MsgBox "終了しました"
            
            End If
            
            
        Case 2
            
            
            
            
            ans = MsgBox("「ＣＳＶ」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                If SAVE_P_SSHIJI_O_Open(BtOpenNomal) Then
                    Unload Me
                End If
                
                
                
                If Output_Proc() Then
                    Unload Me
                End If
            
            
                sts = BTRV(BtOpClose, SAVE_P_SSHIJI_O_POS, SAVE_P_SSHIJI_O_REC, Len(SAVE_P_SSHIJI_O_REC), K0_SAVE_P_SSHIJI_O, Len(K0_SAVE_P_SSHIJI_O), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
                    End If
                End If
            
            
            
                MsgBox "終了しました"
            
            End If
            
            
            


        Case 3
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
    
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
    Check1.Value = vbUnchecked
    Check2.Value = vbUnchecked
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_SHIJI21 = Nothing

    End
End Sub

