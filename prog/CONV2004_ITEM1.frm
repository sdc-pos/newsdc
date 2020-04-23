VERSION 5.00
Begin VB.Form CONV2004_ITEM1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "品目マスタセットアップ処理"
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
   Begin VB.Label Out_Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label In_Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品目ＣＳＶ＝"
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
      Caption         =   "品目マスタセットアップ処理"
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
      Width           =   6240
   End
End
Attribute VB_Name = "CONV2004_ITEM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim In_Count        As Long
Dim Out_Count       As Long

Dim DISP_INTERVAL   As Long

Dim fileName        As String
Dim FileNo          As Integer

Dim c               As String * 128

Dim In_JGYOBU       As String       '事業部
Dim In_NAIGAI       As String       '国内外
Dim In_Hin_Gai      As String       '品番（外部）
Dim In_Packing_No   As String       '個装箱№
Dim In_Jan_Code     As String       'Janコード
Dim In_Hin_Change   As String       '品目読替えコード
Dim In_Goods_Kbn    As String       '商品化有無フラグ
Dim In_EOD          As String

    Update_Proc = True
'---------------------------------------------  品目マスタ追加項目セットアップ
    MsgLab(1) = "品目マスタセットアップ処理中！！"
    Me.MousePointer = vbHourglass
                                                '品目ＣＳＶデータフルパス取込み
    sts = GetIni("FILE", "ITEM_CSV", "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [ITEM_CSV]読み込みエラー ")
        Exit Function
    End If
    fileName = Trim(c)
    
        
    
    On Error GoTo Error_Proc
        
    FileNo = FreeFile
    Open fileName For Input As #FileNo
    
    On Error GoTo 0
    
    
    
    In_Count = 0
    DISP_INTERVAL = 0
    In_Cnt(0).Caption = Format(In_Count, "#0")
                                        
                                        
    Do
        
        DoEvents
            
        On Error GoTo Error_Proc
        Input #FileNo, In_JGYOBU, In_NAIGAI, In_Hin_Gai, In_Packing_No, In_Jan_Code, In_Hin_Change, In_Goods_Kbn, In_EOD
        On Error GoTo 0
        
        In_Count = In_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            In_Cnt(0).Caption = Format(In_Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        '---------------------------    品目マスタ読み込み
        Call UniCode_Conv(K0_ITEM.JGYOBU, In_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, In_NAIGAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, In_Hin_Gai)
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        
        If sts = BtNoErr Then
        
            Call UniCode_Conv(ITEMREC.PACKING_NO, In_Packing_No)    '個装箱番号
            Call UniCode_Conv(ITEMREC.JAN_CODE, In_Jan_Code)        'Janコード
            Call UniCode_Conv(ITEMREC.HIN_CHANGE, In_Hin_Change)    '品番読替えコード
            Call UniCode_Conv(ITEMREC.GOODS_KBN, In_Goods_Kbn)      '商品化有無
        
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
                        Call File_Error(sts, BtOpUpdate, "品目マスタ")
                        Exit Function
                End Select
            Loop
        
        End If
            
    
    Loop

    In_Cnt(0).Caption = Format(In_Count, "#0")

    MsgBox "正常終了しました"
'---------------------------------------------  終了
    Update_Proc = False
    
    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case 62
            In_Cnt(0).Caption = Format(In_Count, "#0")
            MsgBox "正常終了しました"
            Update_Proc = False
            Exit Function
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブが見つかりません" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("ファイルが見つかりません" & fileName, vbExclamation)
        Case 76
            Beep
            ans = MsgBox("ファイルパスが見つかりません" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [ITEM_CSV Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select

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
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_ITEM1 = Nothing

    End
End Sub

