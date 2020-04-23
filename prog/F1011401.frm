VERSION 5.00
Begin VB.Form F1011401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "振替品番 初期設定"
   ClientHeight    =   6936
   ClientLeft      =   2136
   ClientTop       =   2736
   ClientWidth     =   11484
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
   MaxButton       =   0   'False
   OLEDropMode     =   1  '手動
   ScaleHeight     =   6936
   ScaleWidth      =   11484
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      DragMode        =   1  '自動
      Height          =   420
      Index           =   1
      Left            =   2640
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   2
      Top             =   1320
      Width           =   5940
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   420
      Index           =   2
      Left            =   6990
      TabIndex        =   19
      Top             =   5280
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   420
      Index           =   0
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   1
      Top             =   660
      Width           =   720
   End
   Begin VB.ListBox List1 
      Height          =   2928
      ItemData        =   "F1011401.frx":0000
      Left            =   900
      List            =   "F1011401.frx":0002
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2160
      Width           =   10275
   End
   Begin VB.CommandButton Command 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10260
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9420
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8580
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7740
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6420
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5580
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4740
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3900
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2580
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   900
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "登 録"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXCELファイル"
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   22
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label LabJIGYO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   2475
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "登録件数"
      Height          =   255
      Index           =   4
      Left            =   5955
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   990
      TabIndex        =   18
      Top             =   300
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3540
      TabIndex        =   17
      Top             =   660
      Width           =   5025
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "担当者"
      Height          =   255
      Index           =   1
      Left            =   990
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1011401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    '設定用ファイル （CSV)
Dim In_Files    As String   'C:\SDC\FILES\FURIKAE.CSV


Private Const ptxTANTO% = 0
Private Const ptxFILE% = 1
Private Const ptxCNT% = 2

Private Const Text_Max% = 2


Private Const fncDO% = 0
Private Const fncEND% = 11

'Private Const pcmbNAIGAI% = 0

Private Const LAST_UPDATE_DAY$ = "(F101140 2011.06.07 17:00)"


Private Sub List_Proc(DISP_MSG As String)
'----------------------------------------------------------------------------
'                   リストボックス表示処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim W_Edit      As String

        
    W_Edit = ""
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.HIN_MAE, vbUnicode)
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.HIN_GO, vbUnicode)
    
    W_Edit = W_Edit & DISP_MSG & " "
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.BIKOU, vbUnicode)
    
    List1.AddItem W_Edit

End Sub
Private Sub Clear_Field(Mode As Integer)
'----------------------------------------------------------------------------
'                   画面消去処理
'----------------------------------------------------------------------------
Dim i As Integer

    For i = 0 To Text_Max%
        Text1(i).Text = ""
    Next i
    Label1(0).Caption = ""

End Sub
Private Function Error_Check_Proc(index As Integer, Chk_Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case index
    
        Case ptxTANTO%    '担当者
            If Trim(Text1(index)) = "" Then
                MsgBox "担当者未設定エラー"
                Exit Function
            End If
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO%).Text)
    
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            
            Select Case sts
                Case BtNoErr
                    Label1(0).Caption = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                Case BtErrKeyNotFound
                    If Chk_Mode = 0 Then
                        Label1(0).Caption = "担当者コード　未登録"
                    Else
                        MsgBox "入力した項目はエラーです（担当者 未登録）"
                        Text1(index).SetFocus
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
    
        Case ptxFILE%       '設定用ファイル名
            If Trim(Text1(ptxFILE%)) = "" Then
                MsgBox "設定用ファイル名：未設定"
                Text1(index).SetFocus
                Exit Function
            End If
            In_Files = Trim(Text1(ptxFILE%))
            Command(fncDO).Enabled = True
            
        Case Else
    
        
    End Select
        
    Error_Check_Proc = False
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   追加／変更処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim W_MSG           As String
Dim W_CNT           As Long

Dim W_01            As String                      '01項目（振替元品番）
Dim W_02            As String                      '02項目（振替先品番）
Dim W_03            As String                      '03項目（備　考）
Dim W_04            As String                      '10項目（ダミー："*"固定）

Dim W_Book          As Object
Dim X_i             As Long

Dim W_DATE          As String
Dim W_TIME          As String


    Update_Proc = True
    
    W_DATE = Format(Date, "yyyymmdd")
    
    List1.Clear
    
    W_CNT = 0
    
    Call FURIKAE_CLR
    
    Call UniCode_Conv(FURIKAEREC.INS_TANTO, Trim(Text1(ptxTANTO)))
    
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxFILE%).Text) & "]" & "追加処理開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    Set W_Book = GetObject(In_Files)
    
    
'                '設定データＯＰＥＮ(TXT)
'    On Error GoTo Error_Proc
'    Open In_Files For Input As #1 Len = 1024
'
'
'
'            '項目名Rec　Dummy読み
'    Input #1, W_01, W_02, W_03, W_04
'
'
'
'    Do While Not EOF(1)
    X_i = 2
    Do
        W_MSG = ""
        
'        Input #1, W_01, W_02, W_03, W_04
        W_01 = W_Book.Worksheets(1).Range("A" & X_i).Value
        W_02 = W_Book.Worksheets(1).Range("B" & X_i).Value
        W_03 = W_Book.Worksheets(1).Range("C" & X_i).Value
        W_04 = W_Book.Worksheets(1).Range("D" & X_i).Value
        
        If Trim(W_02) = "" Then Exit Do
    
    
        Call UniCode_Conv(K0_FURIKAE.HIN_MAE, Trim(W_02))
        Call UniCode_Conv(K0_FURIKAE.HIN_GO, Trim(W_03))
        Do
            sts = BTRV(BtOpGetGreaterEqual, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Call FURIKAE_CLR
                    'MsgBox "指定された工程がありません。"
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    ans = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                                "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If ans = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "品番振替Ｍ")
                    Exit Function
            End Select
        Loop
        
        If Trim(W_02) <> Trim(StrConv(FURIKAEREC.HIN_MAE, vbUnicode)) Then
            If Trim(W_03) = Trim(StrConv(FURIKAEREC.HIN_GO, vbUnicode)) Then
                W_MSG = "振替品番　登録済み"
            End If
        End If
        
        If W_MSG = "" Then
        
            Call UniCode_Conv(FURIKAEREC.HIN_MAE, Trim(W_02))
            Call UniCode_Conv(FURIKAEREC.HIN_GO, Trim(W_03))
            
            Call UniCode_Conv(FURIKAEREC.BIKOU, Trim(W_04))
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(FURIKAEREC.HIN_MAE, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)              '国内
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Do
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound       'レコード無し
                        
                        'MsgBox "指定された工程がありません。"
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        ans = MsgBox("他で使用中です！<品番Ｍ>" & Chr(13) & Chr(10) & _
                                    "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If ans = vbNo Then Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品番Ｍ")
                        Exit Function
                End Select
            Loop
            
    '        If sts <> BtNoErr Then
    '            W_MSG = "振替元品番　未登録"
    '        End If
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(FURIKAEREC.HIN_GO, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)              '国内
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Do
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound       'レコード無し
                        
                        'MsgBox "指定された工程がありません。"
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        ans = MsgBox("他で使用中です！<品番Ｍ>" & Chr(13) & Chr(10) & _
                                    "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If ans = vbNo Then Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品番Ｍ")
                        Exit Function
                End Select
            Loop
            
    '        If sts <> BtNoErr Then
    '            W_MSG = "振替先品番　未登録"
    '        End If
            
            
            If W_MSG = "" Then
            
                W_TIME = Format(Time, "hhmmss")
                Call UniCode_Conv(FURIKAEREC.Ins_DateTime, W_DATE & W_TIME)
                Call UniCode_Conv(FURIKAEREC.INS_TANTO, Text1(ptxTANTO))
            
                Do
                    sts = BTRV(BtOpInsert, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
                    
                    Select Case sts
                        Case BtNoErr
                            W_CNT = W_CNT + 1
                            Text1(ptxCNT) = Format(W_CNT, "###,###")
                            DoEvents
                            Exit Do
                        Case BtErrDuplicates       'レコード ダブリ
                            W_MSG = "振替品番　登録済み"
                            
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpInsert, "品番振替Ｍ")
                            W_MSG = "登録エラー"
                            Exit Do
                    End Select
                Loop
            End If
        
        End If
        
        If W_MSG <> "" Then
            Call List_Proc(W_MSG)
        End If
        
        
        X_i = X_i + 1
    Loop



'    Close #1

    Text1(ptxCNT) = Format(W_CNT, "###,##0")
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxFILE%).Text) & "]" & "追加処理終了" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    Update_Proc = False
    
    Exit Function
    
Error_Proc:

    If Err.Number = 53 Then
        MsgBox "指定ファイル[ " & In_Files & " ]が、有りません。"
    Else
        MsgBox "Err.Number = " & Err.Number
    End If
    
    On Error Resume Next

End Function


Private Sub Command_Click(index As Integer)

Dim yn      As Integer
Dim sts     As Integer
Dim X_i     As Integer

    Select Case index
        Case fncDO
                                          
            For X_i = ptxTANTO To ptxFILE
                If Error_Check_Proc(X_i, 0) Then    'エラーチェック
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
            Next X_i
                                             
                                             
            yn = MsgBox("登録しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbNo Then
                Command(fncEND).SetFocus
                Exit Sub
            End If
            
            
            Call Input_Lock
            If Update_Proc() Then
                Unload Me
            End If
            Call Input_UnLock
                
            'Call Clear_Field(0)
                
            MsgBox "登録終了しました。"
            
            Command(fncEND).SetFocus
        Case fncEND
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub


Private Sub Form_DblClick()
'    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
        Case vbKeyZ
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

    Show
    DoEvents
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                
                                
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "振替品番ﾏｽﾀﾒﾝﾃﾅﾝｽ" & LAST_UPDATE_DAY, Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1011401.Caption = "振替品番ﾏｽﾀﾒﾝﾃﾅﾝｽ（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
'    Combo1(pcmbNAIGAI).AddItem NAIGAI1$ & " " & NAIGAI_NAI$
'    Combo1(pcmbNAIGAI).AddItem NAIGAI2$ & " " & NAIGAI_GAI$
'    Combo1(pcmbNAIGAI).ListIndex = 0
                                
                                
                                
                                
    Text1(ptxTANTO%).Text = ""
    Text1(ptxFILE%).Text = ""
    
                                '振替品番マスタＯＰＥＮ
    If FURIKAE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '原産国マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
'    If GetIni("F101140", "IN_FILE", "F101140", c) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        End
'    End If
    In_Files = ""

    Text1(ptxFILE%) = In_Files
    
    Command(fncDO).Enabled = False
                                
    Text1(ptxTANTO).SetFocus
    
    End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1(ptxFILE%) = Trim(Data.Files(1))
    
    Text1(ptxFILE%).SetFocus
    
    Call Text1_GotFocus(ptxFILE%)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    Call FURIKAE_CLOSE
        
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品番Ｍ")
        End If
    End If
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者Ｍ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1011401 = Nothing
    End
    
End Sub





Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1011401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011401)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1011401)

    F1011401.MousePointer = vbDefault

End Sub


Private Sub SubMenu_Click(index As Integer)
Dim i As Integer
                                    
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1011401.Caption = "原産国ﾏｽﾀﾒﾝﾃﾅﾝｽ（" + RTrim(JGYOBU_T(index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(index).CODE
    SubMenu(index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(index).COLOR)

End Sub

Private Sub Text1_GotFocus(index As Integer)
    If Text1(index).TabStop = True Then
        Text1(index) = Trim(Text1(index).Text)
        Text1(index).SelStart = 0
        Text1(index).SelLength = Len(Text1(index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(index, 0) Then    'エラーチェック
        Text1(index).SetFocus
        Call Text1_GotFocus(index)
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Text1_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim W_Files     As String

    If index <> ptxFILE% Then Exit Sub
    
    W_Files = Trim(Data.Files(1))
    Text1(ptxFILE%) = Trim(W_Files)
    
    If Trim(Text1(ptxFILE%)) <> "" Then
        Command(fncDO).Enabled = True
    End If
End Sub
