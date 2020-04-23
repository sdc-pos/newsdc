VERSION 5.00
Begin VB.Form F1070501 
   BackColor       =   &H00FFFFFF&
   Caption         =   "部材センター棚卸差異リスト印刷([F107050] 2012.06.26 12:00)"
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
   Begin VB.CheckBox Check1 
      Caption         =   "差異有り分のみ"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   2760
      Value           =   1  'ﾁｪｯｸ
      Width           =   2055
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
      TabIndex        =   13
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
      TabIndex        =   10
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
      TabIndex        =   9
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "対象年月"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1070501"
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



Private Print_JYOGAI_SOKO       As Variant          '印刷除外倉庫(除外倉庫)
Private Print_JYOGAI_SOKO_T()   As String * 2
        
        
Private Print_SHIME_BI          As String * 2       '締め日
Private Print_DATE_S            As String * 8       '使用月範囲　開始
Private Print_DATE_E            As String * 8       '使用月範囲　終了
        


Private Const LMAX% = 44                            '頁内最大行数
Private Const MGN_L% = 10                            '左余白（桁数：１から）
Private Const MGN_U% = 1                            '上余白（行数：１から）

Private Pdate                   As String           '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime                   As String           '印刷開始時刻（ﾍｯﾀﾞｰ用）

Private NormalFont              As New StdFont      '印刷フォント

Private PRT_CAN                 As Boolean          '印刷途中キャンセル要求


Private wkDateTime              As String

Private F107050CSV              As String           'CSV出力ファイル

Private Function Print_Proc() As Integer
'-------------------------------------------------------------------
'
'   棚卸差異リスト　集計　＆　印刷
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Ret             As Integer
Dim FullPath        As String
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer

Dim Sumi_Qty        As Long
Dim Mi_Qty          As Long

Dim Qty             As Long

Dim c               As String * 128

Dim Print_F         As Boolean
Dim RetBuf          As String


    Print_Proc = True

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リストクリアー中", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　棚卸差異リストＦ")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪ＰＣ　棚卸差異リストＦ")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リスト集計中", Me.hwnd, 0)


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

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  棚卸差異F ﾚｺｰﾄﾞ作成
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
        Call UniCode_Conv(K0_ITEM.NAIGAI, "")
        
        com = BtOpGetGreater
                
                
        Do
        
            DoEvents
        
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "棚卸差異リスト印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
        
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                                
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "品目マスタ")
                    Exit Function
            End Select
        
        
        
'2012.04.18            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = "1" Then
        
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, "00000000")
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, "00000000")
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, "00000000")
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.FILLER, "")
        
                sts = BTRV(BtOpInsert, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrDuplicates
                    Case Else
                        Call File_Error(sts, com, "大阪PC　棚卸差異F")
                        Exit Function
                End Select

'2012.04.18            End If

            com = BtOpGetNext
        
        
        Loop
        

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計(部材センター分)
    com = BtOpGetFirst
    
    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リスト印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If

        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.06.26

'        If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, BUZAI, NAIGAI_NAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode)) Then
'            Exit Function
'        End If


        If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, BUZAI, NAIGAI_NAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode), , , , True) Then
            Exit Function
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.06.26



        Qty = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode))
        Qty = Qty + (Sumi_Qty + Mi_Qty)

If Qty <> 0 Then
    Debug.Print
End If


        Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, Format(Qty, "00000000"))

        sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                Exit Function
        End Select
        
        com = BtOpGetNext
            
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計(資材分)
    i = 0
    Do
        
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リスト印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If
        
        
        i = i + 1
        
        GLB_SYUSHI_F = Format(i, "00")
                                            '資材棚卸しﾃﾞｰﾀフルパス取込み
        sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
        If sts <> False Then
            Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]読み込みエラー")
            Exit Function
        End If
        
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
        
        Do
            sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), BtOpenRead)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                    Sleep (500&)
                Case BtErrFileNotFound
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpOpen, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        
        If sts = BtErrFileNotFound Then
            Exit Do
        End If
        
        
        com = BtOpGetFirst
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "棚卸差異リスト印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
        
            sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        
            If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" Then
            Else
                Call UniCode_Conv(K0_OSAKA_TANAOROSHI_SAI.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
                sts = BTRV(BtOpGetEqual, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                            Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "大阪PC　棚卸差異F")
                        Exit Function
                End Select
            
            
            
                Qty = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode))
                
                Qty = Qty + Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, Format(Qty, "00000000"))
        
                sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                            Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                        Exit Function
                End Select
            
            End If
            
            com = BtOpGetNext
        
        Loop
                                                '資材棚卸データ ＣＬＯＳＥ
        sts = BTRV(BtOpClose, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "資材棚卸データ")
            End If
        End If
    
    
    Loop
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  差異の集計
    com = BtOpGetFirst
    
    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リスト印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If

        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select

        Qty = Abs(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) - Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)))

        
                
        Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, Format(Qty, "00000000"))

        sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                Exit Function
        End Select
        
        com = BtOpGetNext
            
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  印刷処理開始
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リスト印刷中", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
    
    Do
    
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リスト印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select
    
        Print_F = True
    
        If Check1.Value = vbChecked Then
            If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, vbUnicode)) = 0 Then
                Print_F = False
            End If
        End If
    
    
        If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) = 0 And _
            Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) = 0 Then
            
            Print_F = False
        
        End If
    
    
    
        If Print_F Then
            'ヘッダーコントロール
            If LCNT > LMAX Then
                Call Print_Head(LCNT)
            End If
    
            '品番
            Printer.Print Tab(MGN_L);
            Printer.Print Left(StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode), 12);
            '品名
            Printer.Print Tab(MGN_L + 21);
            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                                
                Case BtErrKeyNotFound
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            Call Moji_Cut_Proc(StrConv(ITEMREC.HIN_NAME, vbUnicode), RetBuf, 20)
            Printer.Print RetBuf;
            '棚番
            Printer.Print Tab(MGN_L + 50);
            Printer.Print StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_SOKO, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_RETU, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_REN, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_DAN, vbUnicode);
            '資材在庫
            Printer.Print Tab(MGN_L + 70);
            RetBuf = Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            '等号／不等号
            Printer.Print Tab(MGN_L + 85);
            If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) > Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                Printer.Print "＞";
           Else
                If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) < Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                    Printer.Print "＜";
                Else
                    If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                        Printer.Print "＝";
                    End If
                End If
            End If
            '部材センター在庫数
            Printer.Print Tab(MGN_L + 90);
            RetBuf = Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            '差異数
            Printer.Print Tab(MGN_L + 110);
            RetBuf = Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf
            Printer.Print
        
        
            LCNT = LCNT + 2
            
        End If
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
        "棚卸差異リスト印刷終了", Me.hwnd, 0)

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
    Printer.Print "＊＊＊  棚卸差異リスト  ＊＊＊";
    Printer.Print Tab(100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    Printer.Print
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "品  番";
    Printer.Print Tab(MGN_L + 21);
    Printer.Print "　　品             名";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 70);
    Printer.Print "資材在庫";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "部材在庫";
    Printer.Print Tab(MGN_L + 111);
    Printer.Print "差異数"
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1070501.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070501)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070501)


    F1070501.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        Case 7
        
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "入力した項目はエラーです。（対象年月）"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("棚卸差異リスト」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If OutPut_Proc() Then
                    Unload Me
                End If
            End If
        
        
        Case 8                              '印刷
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "入力した項目はエラーです。（対象年月）"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("棚卸差異リスト」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
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
        "棚卸差異リスト印刷", Me.hwnd, 0)
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
        c = "**"
    End If
    Print_Jgyobu = Split(Trim(c), ",", -1)
    Erase Print_Jgyobu_T
        
    For i = 0 To UBound(Print_Jgyobu)
    
        ReDim Preserve Print_Jgyobu_T(0 To i)
        Print_Jgyobu_T(i) = Print_Jgyobu(i)
    Next i
                                
                                
                                
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
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.06.26
                                '印刷除外倉庫
    If GetIni(App.EXEName, "JYOGAI_SOKO", App.EXEName, c) Then
        c = "**"
    End If
    
    
    Zaiko_Syukei_Jyogai_Soko_No = Split(Trim(c), ",", -1)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.06.26

                                
                                'ＣＳＶﾌｧｲﾙ
    If GetIni(App.EXEName, "F107050CSV", App.EXEName, c) Then
    Else
        F107050CSV = Trim(c)
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
                                '大阪PC　棚卸差異F ＯＰＥＮ
    If OSAKA_TANAOROSHI_SAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1070501.FontName
        .Size = F1070501.FontSize
    End With
    Set Printer.Font = NormalFont
    
    Text1(ptxYM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)
    
    Text1(ptxYM).SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
    
    
    
    yn = MsgBox("[棚卸差異リスト印刷]処理を終了しますか？", vbYesNo, "確認入力")
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
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070501 = Nothing

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

Private Function OutPut_Proc() As Integer
'-------------------------------------------------------------------
'
'   棚卸差異リスト　集計　＆　データ出力
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Ret             As Integer
Dim FullPath        As String
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer

Dim Sumi_Qty        As Long
Dim Mi_Qty          As Long

Dim Qty             As Long

Dim c               As String * 128

Dim Print_F         As Boolean
Dim RetBuf          As String

Dim FileNo          As Integer



    OutPut_Proc = True

    Call Input_Lock         '画面項目ロック


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (F107050CSV) For Output As FileNo



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リストクリアー中", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪ＰＣ　棚卸差異リストＦ")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪ＰＣ　棚卸差異リストＦ")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リスト集計中", Me.hwnd, 0)


    wkDateTime = Format(Now, "YYYYMMDDHHMMSS")


'印刷中は「印刷中断」以外のイベント取得不可
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time



    PRT_CAN = False

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  棚卸差異F ﾚｺｰﾄﾞ作成
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
        Call UniCode_Conv(K0_ITEM.NAIGAI, "")
        
        com = BtOpGetGreater
                
                
        Do
        
            DoEvents
        
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "棚卸差異リストデータ出力中断", Me.hwnd, 0)
                Command1.Visible = False
                OutPut_Proc = False
                Exit Function
            End If
        
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                                
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "品目マスタ")
                    Exit Function
            End Select
        
        
        
'2012.04.18            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = "1" Then
        
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, "00000000")
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, "00000000")
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, "00000000")
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.FILLER, "")
        
                sts = BTRV(BtOpInsert, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrDuplicates
                    Case Else
                        Call File_Error(sts, com, "大阪PC　棚卸差異F")
                        Exit Function
                End Select

'2012.04.18            End If

            com = BtOpGetNext
        
        
        Loop
        

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計(部材センター分)
    com = BtOpGetFirst
    
    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リストデータ出力中断", Me.hwnd, 0)
            Command1.Visible = False
            OutPut_Proc = False
            Exit Function
        End If

        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.06.26
'        If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, BUZAI, NAIGAI_NAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode)) Then
'            Exit Function
'        End If

        If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, BUZAI, NAIGAI_NAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode), , , , True) Then
            Exit Function
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.06.26

        Qty = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode))
        Qty = Qty + (Sumi_Qty + Mi_Qty)



        Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, Format(Qty, "00000000"))

        sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                Exit Function
        End Select
        
        com = BtOpGetNext
            
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  在庫情報集計(資材分)
    i = 0
    Do
        
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リストデータ出力中断", Me.hwnd, 0)
            Command1.Visible = False
            OutPut_Proc = False
            Exit Function
        End If
        
        
        i = i + 1
        
        GLB_SYUSHI_F = Format(i, "00")
                                            '資材棚卸しﾃﾞｰﾀフルパス取込み
        sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
        If sts <> False Then
            Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]読み込みエラー")
            Exit Function
        End If
        
        Ret = InStr(1, Trim(c), ".") - 1
        FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
        
        Do
            sts = BTRV(BtOpOpen, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), ByVal FullPath, Len(FullPath), BtOpenRead)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                    Sleep (500&)
                Case BtErrFileNotFound
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpOpen, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        
        If sts = BtErrFileNotFound Then
            Exit Do
        End If
        
        
        com = BtOpGetFirst
    
    
        Do
            DoEvents
                                                '印刷中断要求
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '画面項目ロック解除
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "棚卸差異リスト印刷中断", Me.hwnd, 0)
                Command1.Visible = False
                OutPut_Proc = False
                Exit Function
            End If
        
            sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        
            If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" Then
            Else
                Call UniCode_Conv(K0_OSAKA_TANAOROSHI_SAI.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
                sts = BTRV(BtOpGetEqual, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                            Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "大阪PC　棚卸差異F")
                        Exit Function
                End Select
            
            
            
                Qty = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode))
                
                Qty = Qty + Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
        
                Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, Format(Qty, "00000000"))
        
                sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                            Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                        Exit Function
                End Select
            
            End If
            
            com = BtOpGetNext
        
        Loop
                                                '資材棚卸データ ＣＬＯＳＥ
        sts = BTRV(BtOpClose, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "資材棚卸データ")
            End If
        End If
    
    
    Loop
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  差異の集計
    com = BtOpGetFirst
    
    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リスト印刷中断", Me.hwnd, 0)
            Command1.Visible = False
            OutPut_Proc = False
            Exit Function
        End If

        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select

        Qty = Abs(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) - Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)))

        
                
        Call UniCode_Conv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, Format(Qty, "00000000"))

        sts = BTRV(BtOpUpdate, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "大阪PC　棚卸差異F")
                Exit Function
        End Select
        
        com = BtOpGetNext
            
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  印刷処理開始
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リストデータ出力中", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
    
    Do
    
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "棚卸差異リストデータ出力中断", Me.hwnd, 0)
            Command1.Visible = False
            OutPut_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_TANAOROSHI_SAI_POS, OSAKA_TANAOROSHI_SAI_REC, _
                    Len(OSAKA_TANAOROSHI_SAI_REC), K0_OSAKA_TANAOROSHI_SAI, Len(K0_OSAKA_TANAOROSHI_SAI), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "大阪PC　棚卸差異F")
                Exit Function
        End Select
    
        Print_F = True
    
        If Check1.Value = vbChecked Then
            If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, vbUnicode)) = 0 Then
                Print_F = False
            End If
        End If
    
    
        If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) = 0 And _
            Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) = 0 Then
            
            Print_F = False
        
        End If
    
    
    
        If Print_F Then
            'ヘッダーコントロール
            If LCNT = 99 Then
                
                Write #FileNo, "品  番", "　　品             名", "標準棚番", "資材在庫", "", "部材在庫", "差異数"
                
                LCNT = 0
            End If
    
            '品番
            Write #FileNo, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode),
            '品名
            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                                
                Case BtErrKeyNotFound
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_TANAOROSHI_SAI_REC.HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
            '棚番
            Write #FileNo, StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_SOKO, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_RETU, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_REN, vbUnicode) & "-" & _
                            StrConv(OSAKA_TANAOROSHI_SAI_REC.ST_DAN, vbUnicode),
            '資材在庫
            Write #FileNo, Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)), "#,##0"),
            '等号／不等号
            If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) > Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                Write #FileNo, "＞",
           Else
                If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) < Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                    Write #FileNo, "＜",
                Else
                    If Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SHIZAI_ZAIKO_QTY, vbUnicode)) = Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)) Then
                        Write #FileNo, "＝",
                    End If
                End If
            End If
            '部材センター在庫数
            Write #FileNo, Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.BUZAI_ZAIKO_QTY, vbUnicode)), "#,##0"),
            '差異数
            Write #FileNo, Format(Val(StrConv(OSAKA_TANAOROSHI_SAI_REC.SAI_SU, vbUnicode)), "#,##0")
            
        End If
        com = BtOpGetNext
    
    Loop
                

    Close #FileNo
    MsgBox "「" & F107050CSV & "」は正常に出力されました。"
    
    
    If WriteIni(App.EXEName, "LAST_PRINT_DateTime", App.EXEName, Now) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "LAST_PRINT_DateTime=")
        Unload Me
    End If
    
    
    
    Call Input_UnLock         '画面項目ロック解除
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "棚卸差異リストデータ出力終了", Me.hwnd, 0)

    OutPut_Proc = False

    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox F107050CSV & "が使用中です。"
        OutPut_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

    Call Input_UnLock         '画面項目ロック解除


End Function


