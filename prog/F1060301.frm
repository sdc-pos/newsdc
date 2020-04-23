VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form F1060301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "作業監視モニター（朝一番用）"
   ClientHeight    =   6915
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   12195
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12195
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   9
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   8
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   7
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   5
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5280
      Width           =   732
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
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
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "最　新"
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   4335
      Left            =   1440
      OleObjectBlob   =   "F1060301.frx":0000
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日分"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   30
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   28
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   26
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "分現在"
      Height          =   255
      Index           =   10
      Left            =   10080
      TabIndex        =   17
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "時"
      Height          =   255
      Index           =   9
      Left            =   9240
      TabIndex        =   16
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   15
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   14
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   13
      Top             =   5400
      Width           =   255
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
      TabIndex        =   12
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1060301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxInput_YY = 0           '指定　日付     年
Private Const ptxInput_MM = 1           '指定　日付     月
Private Const ptxInput_DD = 2           '指定　日付　   日


Private Const ptxDATE_YY% = 5           '現在　年
Private Const ptxDATE_MM% = 6           '現在　月
Private Const ptxDATE_DD% = 7           '現在　日
Private Const ptxTIME_HH% = 8           '現在　時
Private Const ptxTIME_MM% = 9           '現在　分

Dim Y_SYUKA     As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 19             '最大列数

Private Const ColSoko_No% = 0           '列 倉庫№
Private Const ColSoko_Name% = 1         '列 倉庫名称
Private Const ColALL_Su% = 2            '列 合計
Private Const ColTUK_Su% = 3            '列 月切り
Private Const ColSPO_Su% = 4            '列 ｽﾎﾟｯﾄ
Private Const ColHJU_Su% = 5            '列 補充
Private Const ColBOU_Su% = 6            '列 貿易

Private Const RowTotal% = 1             '行 合計


Private Function List_Dsp_Proc() As Integer
    
Dim com         As Integer
Dim sts         As Integer
Dim i           As Integer

Dim Row         As Integer
    
Dim Skip_Flg    As Boolean
    
    List_Dsp_Proc = True
    
    Call Input_Lock
                                    
                                    
    For i = ptxInput_YY To ptxInput_DD
        If IsNumeric(Text(i).Text) Then
            If i = ptxInput_YY Then
                Text(i).Text = Format(CLng(Text(i).Text), "0000")
            Else
                Text(i).Text = Format(CLng(Text(i).Text), "00")
            End If
        End If
    Next i
                                    
                                    'テーブルリセット
    Set Y_SYUKA = Nothing
    
    Row = 1
    
    Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    Y_SYUKA(RowTotal, ColSoko_No) = "00"
    Y_SYUKA(RowTotal, ColSoko_Name) = "全倉庫合計"
   
    
    Call UniCode_Conv(K6_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, "")
    Call UniCode_Conv(K6_Y_SYU.HTANABAN, "")
    Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
    Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    
    com = BtOpGetGreaterEqual
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
        Select Case sts
            Case BtNoErr
                                            '事業部ブレーク
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定データ")
                List_Dsp_Proc = SYS_ERR
                Exit Function
        End Select
                                        
                                        
                                        
        Skip_Flg = False

        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) <> Text(ptxInput_YY) & Text(ptxInput_MM) & Text(ptxInput_DD) Then
                                        
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_SPO Or _
                StrConv(Y_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_HJU Then
                Skip_Flg = True
            End If
        End If
                                        
        If Skip_Flg Then
        Else
            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_FIN Then
                                        '合計表示
                Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                    Case CYU_KBN_TUK        '月切り
                        Y_SYUKA(RowTotal, ColALL_Su) = Y_SYUKA(RowTotal, ColALL_Su) + 1
                        Y_SYUKA(RowTotal, ColTUK_Su) = Y_SYUKA(RowTotal, ColTUK_Su) + 1
                    Case CYU_KBN_SPO, CYU_KBN_TOK   'ｽﾎﾟｯﾄ／特売り
                        Y_SYUKA(RowTotal, ColALL_Su) = Y_SYUKA(RowTotal, ColALL_Su) + 1
                        Y_SYUKA(RowTotal, ColSPO_Su) = Y_SYUKA(RowTotal, ColSPO_Su) + 1
                    Case CYU_KBN_HJU                '補充
                        Y_SYUKA(RowTotal, ColALL_Su) = Y_SYUKA(RowTotal, ColALL_Su) + 1
                        Y_SYUKA(RowTotal, ColHJU_Su) = Y_SYUKA(RowTotal, ColHJU_Su) + 1
                    Case CYU_KBN_BOU                '貿易
                        Y_SYUKA(RowTotal, ColALL_Su) = Y_SYUKA(RowTotal, ColALL_Su) + 1
                        Y_SYUKA(RowTotal, ColBOU_Su) = Y_SYUKA(RowTotal, ColBOU_Su) + 1
                    Case Else
                        Debug.Print
                End Select
                                                    
                                                    '標準倉庫設定済み？
                Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                                                '標準倉庫未設定
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, "??")
                    Case Else
                        Call File_Error(sts, com, "倉庫マスタ")
                        Exit Function
                End Select
                    
        
                For i = Min_Row To Y_SYUKA.UpperBound(1)
                    If Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) = Y_SYUKA(i, ColSoko_No) Then
                        Exit For
                    End If
                Next i
            
                If i > Y_SYUKA.UpperBound(1) Then
                    Row = Row + 1
    '絞り込みようが無いので全件表示を目指す
    '                If Row > Max_Row Then
    '                    Beep
    '                    MsgBox "最大表示行数を超えました。"
    '                    Exit Do
    '                End If
                    
                    Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
            
                    Y_SYUKA(Row, ColSoko_No) = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
                    
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            Y_SYUKA(Row, ColSoko_Name) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                        Case BtErrKeyNotFound
                        Case Else
                           Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    End Select
                    i = Row
                End If
            
                
                Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                    Case CYU_KBN_TUK        '月切り
                        Y_SYUKA(i, ColALL_Su) = Y_SYUKA(i, ColALL_Su) + 1
                        Y_SYUKA(i, ColTUK_Su) = Y_SYUKA(i, ColTUK_Su) + 1
                    Case CYU_KBN_SPO, CYU_KBN_TOK   'ｽﾎﾟｯﾄ／特売り
                        Y_SYUKA(i, ColALL_Su) = Y_SYUKA(i, ColALL_Su) + 1
                        Y_SYUKA(i, ColSPO_Su) = Y_SYUKA(i, ColSPO_Su) + 1
                    Case CYU_KBN_HJU                '補充
                        Y_SYUKA(i, ColALL_Su) = Y_SYUKA(i, ColALL_Su) + 1
                        Y_SYUKA(i, ColHJU_Su) = Y_SYUKA(i, ColHJU_Su) + 1
                    Case CYU_KBN_BOU                '貿易
                        Y_SYUKA(i, ColALL_Su) = Y_SYUKA(i, ColALL_Su) + 1
                        Y_SYUKA(i, ColBOU_Su) = Y_SYUKA(i, ColBOU_Su) + 1
                    Case Else
                        Debug.Print StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                End Select
            End If
        End If
        
        com = BtOpGetNext
    
    Loop
    
    Text(ptxDATE_YY).Text = Left(Format(Now, "yyyymmdd"), 4)
    Text(ptxDATE_MM).Text = Mid(Format(Now, "yyyymmdd"), 5, 2)
    Text(ptxDATE_DD).Text = Right(Format(Now, "yyyymmdd"), 2)
    Text(ptxTIME_HH).Text = Left(Format(Now, "HHmmss"), 2)
    Text(ptxTIME_MM).Text = Mid(Format(Now, "HHmmss"), 3, 2)
        
                                    'DBテーブルリンク
    Y_SYUKA.QuickSort Min_Row, (Y_SYUKA.UpperBound(1)), 0, XORDER_ASCEND, XTYPE_STRING
    
    
    
    Set TDBGrid1.Array = Y_SYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    
        
    Call Input_UnLock
    
    List_Dsp_Proc = False
    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060301)


    F1060301.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)

Dim sts As Integer
    
    Select Case Index
        Case 7                              '最新表示
            If List_Dsp_Proc Then           '集計＆表示
                Unload Me
            End If
            Command(7).SetFocus
        
        Case 11                             '終了
            Unload Me
    End Select
    
End Sub


Private Sub Form_Activate()
                                '集計＆表示
    If List_Dsp_Proc Then
        Unload Me
    End If
            
    Command(7).SetFocus

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
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
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1060301.Caption = "作業監視モニター（朝一番用）（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
                                '倉庫マスタOPEN
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データOPEN
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    Text(ptxInput_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxInput_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxInput_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)
    
    
    End Sub



Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060301 = Nothing

    End
End Sub



Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1060301.Caption = "作業監視モニター（朝一番用）（" + RTrim(JGYOBU_T(Index).NAME) + ")"
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
        
    For i = Index + 1 To 2
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i


End Sub
