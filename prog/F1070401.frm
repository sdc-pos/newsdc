VERSION 5.00
Begin VB.Form F1070401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "部材センター循環棚卸表印刷([F107040] 2013.03.25 08:00)"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2250
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   8040
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   7
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   4
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "処理中断"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "終 了"
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印 刷"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "ﾃﾞｰﾀ"
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
      TabIndex        =   14
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   31
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   30
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   29
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "～"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   28
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   27
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   26
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   25
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "棚番範囲"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   24
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "対象年月"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   23
      Top             =   2160
      Width           =   975
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
      TabIndex        =   22
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1070401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxYM% = 0                            '対象年月


Private Const ptxS_SOKO_No% = 1                     '開始   倉庫№
Private Const ptxS_Retu% = 2                        '　     列
Private Const ptxS_Ren% = 3                         '　     連
Private Const ptxS_Dan% = 4                         '　     段
Private Const ptxE_SOKO_No% = 5                     '開始   倉庫№
Private Const ptxE_Retu% = 6                        '　     列
Private Const ptxE_Ren% = 7                         '　     連
Private Const ptxE_Dan% = 8                         '　     段

Private Const Text_Max% = 8                         '画面項目別最大ｲﾝﾃﾞｯｸｽ


Private Print_Jgyobu            As Variant          '印刷対象事業部
Private Print_Jgyobu_T()        As String * 1


Private Print_Yoin_PLUS         As Variant          '印刷対象要因(増加要因)
Private Print_Yoin_PLUS_T()     As String * 2

Private Print_Yoin_MINUS        As Variant          '印刷対象要因(減少要因)
Private Print_Yoin_MINUS_T()    As String * 2

Private Print_JYOGAI_SOKO       As Variant          '印刷除外倉庫(除外倉庫)
Private Print_JYOGAI_SOKO_T()   As String * 2
        
        
Private Print_SHIME_BI          As String * 2       '締め日
Private Print_DATE_S            As String * 8       '使用月範囲　開始
Private Print_DATE_E            As String * 8       '使用月範囲　終了
        


Private Const LMAX% = 44                            '頁内最大行数
Private Const MGN_L% = 3                            '左余白（桁数：１から）
Private Const MGN_U% = 1                            '上余白（行数：１から）

Private Pdate                   As String           '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime                   As String           '印刷開始時刻（ﾍｯﾀﾞｰ用）

Private NormalFont              As New StdFont      '印刷フォント

Private PRT_CAN                 As Boolean          '印刷途中キャンセル要求


Private wkDateTime              As String

Private F107040CSV              As String           'CSV出力ファイル


Private Function Print_Proc() As Integer
'-------------------------------------------------------------------
'
'   棚卸データ　集計　＆　印刷
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean



    Print_Proc = True

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表クリアー中", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表集計中", Me.hwnd, 0)


    wkDateTime = Format(Now, "YYYYMMDDHHMMSS")


'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock         '画面項目ロック
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time



    PRT_CAN = False

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  入出庫情報集計
    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Print_DATE_S)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "循環棚卸表印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Print_DATE_E Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            End Select
    
    
            For j = 0 To UBound(Print_Yoin_PLUS_T)
                If Print_Yoin_PLUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(1) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
    
            For j = 0 To UBound(Print_Yoin_MINUS_T)
                If Print_Yoin_MINUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(2) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
            com = BtOpGetNext
    
        Loop

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計
        
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Text1(ptxS_SOKO_No).Text)
        Call UniCode_Conv(K0_ZAIKO.Retu, Text1(ptxS_Retu).Text)
        Call UniCode_Conv(K0_ZAIKO.Ren, Text1(ptxS_Ren).Text)
        Call UniCode_Conv(K0_ZAIKO.Dan, Text1(ptxS_Dan).Text)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Text1(ptxS_Dan).Text)
    
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "循環棚卸表印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select
    
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = Print_Jgyobu_T(i) Then
                If Update_Proc(3) Then
                    Exit Function
                End If
            End If
    
            com = BtOpGetNext
                
        Loop
                
    Next i
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  印刷処理開始
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表印刷中", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
                
                
                
    Do
    
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "循環棚卸表印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
    
    
        'ヘッダーコントロール
        If LCNT > LMAX Then
            Call Print_Head(LCNT)
        End If
    
    
        '事業部
        For j = 0 To UBound(JGYOBU_T)
            If JGYOBU_T(j).CODE = StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode) Then
                Exit For
            End If
        Next j
        Printer.Print Tab(MGN_L);
        If j <= UBound(JGYOBU_T) Then
            Call Moji_Cut_Proc(JGYOBU_T(j).NAME, RetBuf, 10)
            Printer.Print RetBuf;
        End If
        '棚番
        Printer.Print Tab(MGN_L + 15);
        Printer.Print StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode);
        '品番
        Printer.Print Tab(MGN_L + 30);
        Printer.Print Left(StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode), 12);
        '品名
        Printer.Print Tab(MGN_L + 45);
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OSAKA_PSTOCKREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        Call Moji_Cut_Proc(StrConv(ITEMREC.HIN_NAME, vbUnicode), RetBuf, 20)
        Printer.Print RetBuf;
        '入庫数
        Printer.Print Tab(MGN_L + 77);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;
        '出庫数
        Printer.Print Tab(MGN_L + 87);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;
        '在庫数
        Printer.Print Tab(MGN_L + 97);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf
        Printer.Print
        
        
        LCNT = LCNT + 2
        
        com = BtOpGetNext
    
    Loop
                

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    
    If WriteIni(App.EXEName, "LAST_PRINT_DateTime", App.EXEName, Now) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "LAST_PRINT_DateTime=")
        Unload Me
    End If
    
    
    
    Call Input_UnLock         '画面項目ロック解除
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表印刷終了", Me.hwnd, 0)

    Print_Proc = False
End Function

Private Sub Print_Head(LCNT As Integer)
                                        
Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    Printer.Print Text1(ptxYM).Text; " 月分"
    Printer.Print Tab(36);
    Printer.Print "＊＊＊  「部材センター」循環棚卸表  ＊＊＊";
    Printer.Print Tab(100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    Printer.Print
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "事業部";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "棚    番";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "品  番";
    Printer.Print Tab(MGN_L + 41);
    Printer.Print "　　品             名";
    Printer.Print Tab(MGN_L + 80);
    Printer.Print "増加分";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "減少分";
    Printer.Print Tab(MGN_L + 100);
    Printer.Print "現在庫"
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1070401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070401)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070401)


    F1070401.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ＣＳＶ出力  2012.04.19
        Case 7
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "入力した項目はエラーです。（対象年月）"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_SOKO_No).Text) = "" Then
                Text1(ptxE_SOKO_No).Text = "zz"
            End If
            If Trim(Text1(ptxE_Retu).Text) = "" Then
                Text1(ptxE_Retu).Text = "zz"
            End If
            If Trim(Text1(ptxE_Ren).Text) = "" Then
                Text1(ptxE_Ren).Text = "zz"
            End If
            If Trim(Text1(ptxE_Dan).Text) = "" Then
                Text1(ptxE_Dan).Text = "zz"
            End If
            If (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) > _
                (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
                MsgBox "入力した項目はエラーです。（棚番範囲）"
                Text1(ptxS_SOKO_No).SetFocus
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「循環棚卸表」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Output_Proc() Then
                    Unload Me
                End If
            End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ＣＳＶ出力
        
        
        
        
        
        
        
        
        
        
        
        Case 8                              '印刷
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "入力した項目はエラーです。（対象年月）"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_SOKO_No).Text) = "" Then
                Text1(ptxE_SOKO_No).Text = "zz"
            End If
            If Trim(Text1(ptxE_Retu).Text) = "" Then
                Text1(ptxE_Retu).Text = "zz"
            End If
            If Trim(Text1(ptxE_Ren).Text) = "" Then
                Text1(ptxE_Ren).Text = "zz"
            End If
            If Trim(Text1(ptxE_Dan).Text) = "" Then
                Text1(ptxE_Dan).Text = "zz"
            End If
            If (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) > _
                (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
                MsgBox "入力した項目はエラーです。（棚番範囲）"
                Text1(ptxS_SOKO_No).SetFocus
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「循環棚卸表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
            End If
                    
        Case 11                             '終了
            Unload Me
    End Select
End Sub
Private Sub Command1_Click()
    PRT_CAN = True
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

Dim wkYY        As Integer
Dim wkMM        As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    

    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表印刷 ", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = Trim(c)
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                
                                '印刷対象事業部
    If GetIni(App.EXEName, "JGYOBU_CODE", App.EXEName, c) Then
        MsgBox "印刷対象事業部の獲得に失敗しました(JGYOBU_CODE=)。処理を中止します。"
        End
    Else
        Print_Jgyobu = Split(Trim(c), ",", -1)
        Erase Print_Jgyobu_T
        
        For i = 0 To UBound(Print_Jgyobu)
        
            ReDim Preserve Print_Jgyobu_T(0 To i)
            Print_Jgyobu_T(i) = Print_Jgyobu(i)
        Next i
        
        
    End If
                                '締め日
    If GetIni(App.EXEName, "SHIME_BI", App.EXEName, c) Then
        MsgBox "締め日の獲得に失敗しました(SHIME_BI=)。処理を中止します。"
        End
    Else
        Print_SHIME_BI = Trim(c)
        If Not IsNumeric(Print_SHIME_BI) Then
            MsgBox "締め日の獲得に失敗しました(SHIME_BI=)。処理を中止します。"
            End
        End If
        
        
        '開始日
        If Mid(Format(Date, "YYYYMMDD"), 7, 2) > Print_SHIME_BI Then
            Print_DATE_S = Mid(Format(Date, "YYYYMMDD"), 1, 6) & Format(Val(Print_SHIME_BI) + 1, "00")
        Else
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
            wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) - 1
            If wkMM < 1 Then
                wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) - 1
                wkMM = 12
            End If
            Print_DATE_S = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(Print_SHIME_BI) + 1, "00")
        End If
        '終了日
        If Mid(Format(Date, "YYYYMMDD"), 7, 2) <= Print_SHIME_BI Then
            Print_DATE_E = Mid(Format(Date, "YYYYMMDD"), 1, 6) & Print_SHIME_BI
        Else
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
            wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) + 1
            If wkMM > 12 Then
                wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) + 1
                wkMM = 1
            End If
            Print_DATE_E = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(Print_SHIME_BI) + 1, "00")
        End If
    End If
                                
                                '印刷対象要因（増加）
    If GetIni(App.EXEName, "YOIN_CODE_PLUS", App.EXEName, c) Then
        c = "**"
    End If
    Print_Yoin_PLUS = Split(Trim(c), ",", -1)
    Erase Print_Yoin_PLUS_T
    For i = 0 To UBound(Print_Yoin_PLUS)
    
        ReDim Preserve Print_Yoin_PLUS_T(0 To i)
        Print_Yoin_PLUS_T(i) = Print_Yoin_PLUS(i)
    Next i
                                
                                
                                
                                '印刷対象要因（減少）
    If GetIni(App.EXEName, "YOIN_CODE_MINUS", App.EXEName, c) Then
        c = "**"
    End If
    Print_Yoin_MINUS = Split(Trim(c), ",", -1)
    
    Erase Print_Yoin_MINUS_T
        
    For i = 0 To UBound(Print_Yoin_MINUS)
    
        ReDim Preserve Print_Yoin_MINUS_T(0 To i)
        Print_Yoin_MINUS_T(i) = Print_Yoin_MINUS(i)
    Next i
                                '印刷除外倉庫
    If GetIni(App.EXEName, "JYOGAI_SOKO", App.EXEName, c) Then
        c = "**"
    End If
    
    
    Print_JYOGAI_SOKO = Split(Trim(c), ",", -1)
    Erase Print_JYOGAI_SOKO_T
        
    For i = 0 To UBound(Print_JYOGAI_SOKO)
    
        ReDim Preserve Print_JYOGAI_SOKO_T(0 To i)
        Print_JYOGAI_SOKO_T(i) = Print_JYOGAI_SOKO(i)
    Next i
                                
                                
                                'ＣＳＶﾌｧｲﾙ
    If GetIni(App.EXEName, "F107040CSV", App.EXEName, c) Then
    Else
        F107040CSV = Trim(c)
        Command(7).Enabled = True
    End If
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '大阪ＰＣ　循環棚卸ＦＯＰＥＮ
    If OSAKA_PSTOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1070401.FontName
        .Size = F1070401.FontSize
    End With
    Set Printer.Font = NormalFont
    
    Text1(ptxYM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)
    
    Text1(ptxYM).SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
    
    
    
    yn = MsgBox("[循環棚卸表印刷]処理を終了しますか？", vbYesNo, "確認入力")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '大阪ＰＣ　循環棚卸ＦＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "大阪ＰＣ　循環棚卸Ｆ")
        End If
    End If
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070401 = Nothing

    End
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text1(i).Enabled And Text1(i).Visible And Text1(i).TabStop Then
            Text1(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Update_Proc(Mode As Integer) As Integer
'-------------------------------------------------------------------
'
'   棚卸データ　作成
'
'   mode :  1:入庫更新
'           2:出庫更新
'           3:在庫更新
'
'-------------------------------------------------------------------
Dim sts         As Integer

Dim wkSoko      As String * 2
Dim wkRetu      As String * 2
Dim wkRen       As String * 2
Dim wkDan       As String * 2


Dim wkJGYOBU    As String * 1
Dim wkNaigai    As String * 1
Dim wkHin_GAI   As String * 20

Dim wkQTY       As Long

Dim com         As Integer
Dim i           As Integer

    Update_Proc = True


    Select Case Mode
        Case 1
            If (StrConv(IDOREC.TO_SOKO, vbUnicode) & StrConv(IDOREC.TO_RETU, vbUnicode) & StrConv(IDOREC.TO_REN, vbUnicode) & StrConv(IDOREC.TO_DAN, vbUnicode)) < (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) Or _
                (StrConv(IDOREC.TO_SOKO, vbUnicode) & StrConv(IDOREC.TO_RETU, vbUnicode) & StrConv(IDOREC.TO_REN, vbUnicode) & StrConv(IDOREC.TO_DAN, vbUnicode)) > (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
    
                Update_Proc = False
                Exit Function
    
            Else
            
                wkSoko = StrConv(IDOREC.TO_SOKO, vbUnicode)
                wkRetu = StrConv(IDOREC.TO_RETU, vbUnicode)
                wkRen = StrConv(IDOREC.TO_REN, vbUnicode)
                wkDan = StrConv(IDOREC.TO_DAN, vbUnicode)
            
                wkJGYOBU = StrConv(IDOREC.JGYOBU, vbUnicode)
                wkNaigai = StrConv(IDOREC.NAIGAI, vbUnicode)
                wkHin_GAI = StrConv(IDOREC.HIN_GAI, vbUnicode)
            
            End If
    
        Case 2

            If (StrConv(IDOREC.FROM_SOKO, vbUnicode) & StrConv(IDOREC.FROM_RETU, vbUnicode) & StrConv(IDOREC.FROM_REN, vbUnicode) & StrConv(IDOREC.FROM_DAN, vbUnicode)) < (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) Or _
                (StrConv(IDOREC.FROM_SOKO, vbUnicode) & StrConv(IDOREC.FROM_RETU, vbUnicode) & StrConv(IDOREC.FROM_REN, vbUnicode) & StrConv(IDOREC.FROM_DAN, vbUnicode)) > (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
    
                Update_Proc = False
                Exit Function
    
            Else
            
                wkSoko = StrConv(IDOREC.FROM_SOKO, vbUnicode)
                wkRetu = StrConv(IDOREC.FROM_RETU, vbUnicode)
                wkRen = StrConv(IDOREC.FROM_REN, vbUnicode)
                wkDan = StrConv(IDOREC.FROM_DAN, vbUnicode)
    
    
                wkJGYOBU = StrConv(IDOREC.JGYOBU, vbUnicode)
                wkNaigai = StrConv(IDOREC.NAIGAI, vbUnicode)
                wkHin_GAI = StrConv(IDOREC.HIN_GAI, vbUnicode)
    
    
            End If
    
    
        Case 3
    
            wkSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            wkRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
            wkRen = StrConv(ZAIKOREC.Ren, vbUnicode)
            wkDan = StrConv(ZAIKOREC.Dan, vbUnicode)
    
            wkJGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode)
            wkNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            wkHin_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
    
    End Select

    
    For i = 0 To UBound(Print_JYOGAI_SOKO_T)
    
        If wkSoko = Print_JYOGAI_SOKO_T(i) Then
    
            Update_Proc = False
            Exit Function
        
        End If
    
    Next i
    
    
    
'-------------------------------------------------------------  大阪ＰＣ　循環棚卸Ｆ
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Soko_No, wkSoko)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Retu, wkRetu)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Ren, wkRen)
    
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Dan, wkDan)               '2013.03.23
    
    Call UniCode_Conv(K0_OSAKA_PSTOCK.JGYOBU, wkJGYOBU)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.NAIGAI, wkNaigai)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.HIN_GAI, wkHin_GAI)
    
    sts = BTRV(BtOpGetEqual, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
    Select Case sts
        Case BtNoErr
        
            com = BtOpUpdate
        
        Case BtErrKeyNotFound
        
            com = BtOpInsert
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "大阪ＰＣ　循環棚卸Ｆ")
            Exit Function
    End Select
    
    If com = BtOpInsert Then
        
        Call UniCode_Conv(OSAKA_PSTOCKREC.Soko_No, wkSoko)      '倉庫№
        Call UniCode_Conv(OSAKA_PSTOCKREC.Retu, wkRetu)         '棚番　列
        Call UniCode_Conv(OSAKA_PSTOCKREC.Ren, wkRen)           '棚番　連
        Call UniCode_Conv(OSAKA_PSTOCKREC.Dan, wkDan)           '棚番　段
        Call UniCode_Conv(OSAKA_PSTOCKREC.JGYOBU, wkJGYOBU)     '事業部区分
        Call UniCode_Conv(OSAKA_PSTOCKREC.NAIGAI, wkNaigai)     '国内外
        Call UniCode_Conv(OSAKA_PSTOCKREC.HIN_GAI, wkHin_GAI)   '品番（外部）
                                                                '計上年月
        Call UniCode_Conv(OSAKA_PSTOCKREC.KEIJYO_YM, Left(Format(Text1(ptxYM).Text & "/01"), 7))
                                                                '当月入庫数
        Call UniCode_Conv(OSAKA_PSTOCKREC.NYUKO_QTY, "0000000000")
                                                                '当月出庫数
        Call UniCode_Conv(OSAKA_PSTOCKREC.SYUKO_QTY, "0000000000")
                                                                '当月在庫残数
        Call UniCode_Conv(OSAKA_PSTOCKREC.ZAIKO_QTY, "0000000000")
        
        Call UniCode_Conv(OSAKA_PSTOCKREC.FILLER, "")
                                                                'ﾃﾞｰﾀ作成日時
        Call UniCode_Conv(OSAKA_PSTOCKREC.Ins_DateTime, wkDateTime)
    
    
    End If
    
    
    Select Case Mode
        Case 1
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode))
            wkQTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.NYUKO_QTY, Format(wkQTY, "0000000000"))
        Case 2
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode))
            wkQTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.SYUKO_QTY, Format(wkQTY, "0000000000"))
        Case 3
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode))
            wkQTY = wkQTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.ZAIKO_QTY, Format(wkQTY, "0000000000"))
    End Select


    sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrDuplicates
        
            Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ" & "KEY =" & StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode) & StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
                    
        Case Else
            Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ")
            Exit Function
    End Select

    Update_Proc = False


End Function
Private Function Output_Proc() As Integer
'-------------------------------------------------------------------
'
'   棚卸データ　集計　＆　データ出力
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean

Dim FileNo          As Integer


    Output_Proc = True


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (F107040CSV) For Output As FileNo


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表クリアー中", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表集計中", Me.hwnd, 0)


    wkDateTime = Format(Now, "YYYYMMDDHHMMSS")


'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock         '画面項目ロック
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time



    PRT_CAN = False

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  入出庫情報集計
    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Print_DATE_S)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "循環棚卸表データ出力中断", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Print_DATE_E Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            End Select
    
    
            For j = 0 To UBound(Print_Yoin_PLUS_T)
                If Print_Yoin_PLUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(1) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
    
            For j = 0 To UBound(Print_Yoin_MINUS_T)
                If Print_Yoin_MINUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(2) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
            com = BtOpGetNext
    
        Loop

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計
        
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Text1(ptxS_SOKO_No).Text)
        Call UniCode_Conv(K0_ZAIKO.Retu, Text1(ptxS_Retu).Text)
        Call UniCode_Conv(K0_ZAIKO.Ren, Text1(ptxS_Ren).Text)
        Call UniCode_Conv(K0_ZAIKO.Dan, Text1(ptxS_Dan).Text)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Text1(ptxS_Dan).Text)
    
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "循環棚卸表データ出力中断", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select
    
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = Print_Jgyobu_T(i) Then
                If Update_Proc(3) Then
                    Exit Function
                End If
            End If
    
            com = BtOpGetNext
                
        Loop
                
    Next i
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  印刷処理開始
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表データ出力中", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
                
                
                
    Do
    
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "循環棚卸表データ出力中断", Me.hwnd, 0)
            Command1.Visible = False
            Output_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　循環棚卸Ｆ")
                Exit Function
        End Select
    
    
        'ヘッダーコントロール
        If LCNT = 99 Then
                
            Write #FileNo, "事業部", "棚    番", "品  番", "　　品             名", "増加分", "減少分", "現在庫"
   
            LCNT = 0
        End If
    
    
        '事業部
        For j = 0 To UBound(JGYOBU_T)
            If JGYOBU_T(j).CODE = StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode) Then
                Write #FileNo, JGYOBU_T(j).NAME,
                Exit For
            End If
        Next j
        '棚番
        Write #FileNo, StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode),
        '品番
        Write #FileNo, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode),
        '品名
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OSAKA_PSTOCKREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
        Printer.Print RetBuf;
        '入庫数
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode)), "#,##0"),
        '出庫数
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode)), "#,##0"),
        '在庫数
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode)), "#,##0")
        
        
        
        com = BtOpGetNext
    
    Loop
                
    Close #FileNo
    MsgBox "「" & F107040CSV & "」は正常に出力されました。"
    
    
    If WriteIni(App.EXEName, "LAST_PRINT_DateTime", App.EXEName, Now) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "LAST_PRINT_DateTime=")
        Unload Me
    End If
    
    
    
    Call Input_UnLock         '画面項目ロック解除
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "循環棚卸表データ終了", Me.hwnd, 0)

    Output_Proc = False

    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox F107040CSV & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

End Function


