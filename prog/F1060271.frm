VERSION 5.00
Begin VB.Form F1060271 
   BackColor       =   &H00FFFFFF&
   Caption         =   "商品化計画支援アラームリスト(出荷ﾃﾞｰﾀ対応)印刷"
   ClientHeight    =   6948
   ClientLeft      =   2328
   ClientTop       =   2712
   ClientWidth     =   11292
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
   ScaleHeight     =   6948
   ScaleWidth      =   11292
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   5670
      MaxLength       =   2
      TabIndex        =   20
      Top             =   1920
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   18
      Top             =   1920
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4515
      MaxLength       =   4
      TabIndex        =   16
      Top             =   1920
      Width           =   540
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
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
         Size            =   11.4
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
         Size            =   11.4
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
      Caption         =   "印 刷"
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
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   5565
      TabIndex        =   19
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   1
      Left            =   5145
      TabIndex        =   17
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "対象日付"
      Height          =   255
      Index           =   0
      Left            =   3465
      TabIndex        =   15
      Top             =   2040
      Width           =   1050
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
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
Attribute VB_Name = "F1060271"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxYY% = 0                    '指定日付　年
Private Const ptxMM% = 1                    '指定日付　年
Private Const ptxDD% = 2                    '指定日付　年

Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNaigai% = 0               '国内外


Private Const LMAX% = 36                    '頁内最大行数
Private Const LCTL% = 99                    '
Private Const MGN_L% = 10                   '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private NormalFont  As New StdFont          '印刷フォント
Private MidFont     As New StdFont          '印刷フォント

Private OutSide     As Long                 '印刷対外出荷数

Private GOODS_DATA  As String               '出力データファイル名

Private NON_MTS     As String               '除外向け先


Private Type EE_ZAIKO_TBL_tag
    EE_LOC          As String * 8
    EE_QTY          As Long
End Type

Private EE_ZAIKO_TBL(0 To 2) As EE_ZAIKO_TBL_tag

Private SHO_SOKO    As Variant              '商品化用倉庫(未商品とみなす分)

Private Const Last_Update_day$ = "([F106027] 2011.07.14 12:00)"

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   エラーチェック処理
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
            
    If Trim(Text(ptxYY).Text) = "" Then
        Text(ptxYY).Text = ""
        Text(ptxMM).Text = ""
        Text(ptxDD).Text = ""
        Err_Chk = False
        Exit Function
    End If
            
    For i = ptxYY To ptxDD
        
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(i).SetFocus
            Exit Function
        
        End If
    
        If i <> ptxYY Then
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    
    Next i
            
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060271.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060271)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060271)


    F1060271.MousePointer = vbDefault

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
    Select Case Index
        
        Case 7                              'データ出力
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                If Data_Proc() Then
                    Unload Me
                End If
            End If
            
            Text(ptxYY).SetFocus
        
        
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxYY).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
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

'
Private Sub Form_Load()

Dim c   As String * 128
Dim i   As Integer
     
     If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
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

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1060271.Caption = "商品化計画支援アラームリスト(出荷ﾃﾞｰﾀ対応)印刷（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '商品化支援ファイル名取り込み
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'商品化支援ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    GOODS_DATA = Trim(c)
'-----------    SYS.INI ---> (自ﾌﾟﾛｸﾞﾗﾑID).INI 2011.07.14
                                '対象外出荷数取り込み
    If GetIni(App.EXEName, "OUTSIDE", App.EXEName, c) Then
        OutSide = 0
    Else
        If IsNumeric(Trim(c)) Then
            OutSide = CLng(Trim(c))
        Else
            OutSide = 0
        End If
    End If
                                '商品化用倉庫取り込み
    If GetIni(App.EXEName, "SHO_SOKO", App.EXEName, c) Then
        c = " "
    End If
    SHO_SOKO = Split(Trim(c), ",", -1)
'-----------    SYS.INI ---> (自ﾌﾟﾛｸﾞﾗﾑID).INI 2011.07.14
                                
                                '除外向け先取り込み
    If GetIni("PI00010", "MTSSS", "P_SYS", c) Then
        NON_MTS = ""
    Else
        NON_MTS = Trim(c)
    End If
                                
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷データファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化集計ファイルＯＰＥＮ
    If GOODS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定(通常)
    With NormalFont
        .NAME = F1060271.FontName
        .Size = 12
    End With

                                '印刷フォント設定（小）
    With MidFont
        .NAME = F1060271.FontName
        .Size = 8
    End With


    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    Text(ptxYY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxMM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxDD).Text = Right(Format(Now, "YYYYMMDD"), 2)


    Show
    
    Text(ptxYY).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
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
                                            '出荷データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷データ")
        End If
    End If
                                            '月平均出荷数ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷数")
        End If
    End If
                                            '商品化集計ファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化集計ファイル")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060271 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1060271.Caption = "商品化計画支援アラームリスト(出荷ﾃﾞｰﾀ対応)印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）" & Last_Update_day
    Last_JGYOBU = JGYOBU_T(Index).CODE
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
Dim sts As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case ptxYY
            If Trim(Text(i).Text) = "" Then
                Text(ptxMM).Text = ""
                Text(ptxDD).Text = ""
                Exit Sub
            End If
    
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(i).SetFocus
                Exit Sub
            End If
        Case ptxMM, ptxDD
    
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(i).SetFocus
                Exit Sub
            Else
                Text(i).Text = Format(CInt(Text(i).Text), "00")
            End If
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   商品化支援アラームリスト印刷処理
'----------------------------------------------------------------------------
Dim LCNT        As Integer

Dim sts         As Integer
Dim com         As Integer
Dim FSW         As Boolean

Dim Save_Soko   As String * 2

Dim Edit        As String



Dim X_Tab       As Integer

    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    End If



    LCNT = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    Call UniCode_Conv(K1_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K1_GOODS.ST_RETU, "")
    Call UniCode_Conv(K1_GOODS.ST_REN, "")
    Call UniCode_Conv(K1_GOODS.ST_DAN, "")
    Call UniCode_Conv(K1_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K1_GOODS.HIN_GAI, "")
    
    com = BtOpGetGreater
    FSW = True
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select

If Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) = "EHA4402057H" Then
    Debug.Print
End If

'-------------------------------------------------  明細印刷
        
        
        
'        If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
'                        '設定発注点より大きい
'            Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
'            Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzz")
'            com = BtOpGetGreater
'        Else
            '未商品在庫＝０ は、印刷対象外 2004.08.27
            If CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)) <= 0 Or _
                CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
            Else
                
                If FSW Then
                    Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                    
                    Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                                Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                            End If
                        Case BtErrKeyNotFound
                            '考えられないが処理は継続
                            Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    End Select
                    FSW = False
                End If
                
                
                
                If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                    
                    LCNT = LMAX + 1
                    Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                    
                    Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                                Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                            End If
                        
                        Case BtErrKeyNotFound
                                '考えられないが処理は継続
                            Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    End Select
                    
                End If
                
                
                
                
                
                If Head_Print_Proc(LCNT) Then
                    Exit Function
                End If
            
                X_Tab = MGN_L
            
                Printer.Print Tab(X_Tab);
                                                        '標準棚番
                Edit = StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 5
'                X_Tab = X_Tab + Len(Edit) + 3
                                                        '品番（外部）
                Printer.Print Tab(X_Tab);

                Printer.Print Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 5
                X_Tab = X_Tab + Len(Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13)) + 4
                                                        '箱№
                Printer.Print Tab(X_Tab);
                Printer.Print StrConv(GOODSREC.PACKING_NO, vbUnicode);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 5
                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                                                        '商品化済み在庫数
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '未商品在庫数
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '月平均出荷数
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '事前商品化必要数
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '事前商品化状況
                Printer.Print Tab(X_Tab);
                Edit = Format(CInt(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If

                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 5
                                                        '別置在庫
                Printer.Print Tab(X_Tab);

                If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                End If

                Edit = ""
                If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
                    Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#,##0")
                    If Len(Edit) < 9 Then
                        Edit = Space(9 - Len(Edit)) & Edit
                    End If
                    Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & _
                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & _
                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & _
                           Right(EE_ZAIKO_TBL(0).EE_LOC, 2) & Edit
                End If

                Printer.Print Edit

                Printer.Print
            
                LCNT = LCNT + 2
        
            End If
            com = BtOpGetNext
        'End If
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(LCNT As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If LCNT < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    If LCNT = LCTL Then
    Else
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i

    Printer.Print Tab(MGN_L + 35);
    
    Printer.Print "商品化支援アラームリスト(出荷データ対応)";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫：";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  "
'    Printer.Print "（設定発注点 " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "％）"
    Printer.Print

'    Printer.Print Tab(MGN_L);
'    Printer.Print "標準棚番";
'    Printer.Print Tab(MGN_L + 13);
'    Printer.Print "品番（外部）";
'    Printer.Print Tab(MGN_L + 26);
'    Printer.Print "資材(箱№)";
'    Printer.Print Tab(MGN_L + 38);
'    Printer.Print "商品化済在庫";
'    Printer.Print Tab(MGN_L + 58);
'    Printer.Print "未商品在庫";
'    Printer.Print Tab(MGN_L + 74);
'    Printer.Print "月平均出荷数";
'    Printer.Print Tab(MGN_L + 88);
'    Printer.Print "事前商品化必要数";
'    Printer.Print Tab(MGN_L + 108);
'    Printer.Print "事前商品化状況"
'
'    Set Printer.Font = MidFont
'    Printer.Print Tab(MGN_L + 112);
'    Printer.Print "(過去3ｹ月間平均)";
'    Printer.Print Tab(MGN_L + 130);
'    Printer.Print "(月平均出荷数-商品化済在庫)";
'    Printer.Print Tab(MGN_L + 158);
'    Printer.Print "(商品化済在庫/月平均出荷数)"
'
'
'    Set Printer.Font = NormalFont

    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "資材";
    Printer.Print Tab(MGN_L + 42);
    Printer.Print "商済数";
    Printer.Print Tab(MGN_L + 54);
    Printer.Print "未商品";
    Printer.Print Tab(MGN_L + 62);
    Printer.Print "本日出荷数";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "必要数";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "　状況";
    Printer.Print Tab(MGN_L + 113);
    Printer.Print "別置在庫"

    Printer.Print

    LCNT = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   支援用集計データ作成処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer


Dim Skip_Flg    As Boolean
Dim Save_Hin_No As String
    
Dim Syuka_Qty   As Long
    
    Data_Make_Proc = True

'---------------------------------------------------------- '全レコード削除
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<GOODS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "商品化支援集計データ")
                    Exit Function
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            
            sts = BTRV(BtOpDelete, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<GOODS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "商品化支援集計データ")
                    Exit Function
            End Select
        
        Loop
        
        com = BtOpGetNext
    
    Loop
'---------------------------------------------------------- '出荷予定ベースでデータ作成

    Save_Hin_No = ""
    
    Call UniCode_Conv(K7_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K7_Y_SYU.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K7_Y_SYU.KEY_HIN_NO, "")
    Call UniCode_Conv(K7_Y_SYU.KEY_SYUKA_YMD, "")
    
    
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K7_Y_SYU, Len(K7_Y_SYU), 7)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    '事業部／国内外ブレーク
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select
        
        
        Skip_Flg = False
        If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_UN Then
            Skip_Flg = True
        End If
        
        If Trim(Text(ptxYY).Text) <> "" Then
            If (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text) <> StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
        
        
        If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) & Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) = NON_MTS Then
            Skip_Flg = True
        End If
        
        If Not Skip_Flg Then
            
            
            If Save_Hin_No = "" Then
                Save_Hin_No = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                Syuka_Qty = 0
            End If
            
            If Save_Hin_No <> StrConv(Y_SYUREC.HIN_NO, vbUnicode) Then
            
                If Data_Make_Sub(Save_Hin_No, Syuka_Qty) Then
                    Exit Function
                End If
            
                Save_Hin_No = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                Syuka_Qty = 0
            
            
            End If
        
        
            Syuka_Qty = Syuka_Qty + CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
        
        End If
        
        com = BtOpGetNext
    
    Loop

    If Save_Hin_No <> "" Then
    
        If Data_Make_Sub(Save_Hin_No, Syuka_Qty) Then
            Exit Function
        End If
    
    End If

    Data_Make_Proc = False


End Function

Private Function Data_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＣＳＶデータ作成処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Save_Soko       As String * 2

Dim Edit            As String

Dim FileNo          As Integer
Dim fileName        As String
    
    
    Data_Proc = True

    Call Input_Lock

    fileName = GOODS_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - sts)
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo


    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    End If
    
    Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K0_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select
'-------------------------------------------------  明細印刷
        
        If com = BtOpGetGreater Then
            Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
            
            Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    End If
                Case BtErrKeyNotFound
                    '考えられないが処理は継続
                    Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Exit Function
            End Select
                    'ヘッダー出力
            Write #FileNo, "*** 商品化支援アラームリスト　***",
            Write #FileNo, "作成日付:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                    
        
            Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況"
            
        
            Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
            
        
        
        End If
        
        If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                            
            Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
            
            Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    End If
                
                Case BtErrKeyNotFound
                        '考えられないが処理は継続
                    Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Exit Function
            End Select
            
            Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
            
            
        End If
        
        
        If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                        '設定発注点より大きい
            Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
            Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzz")
            com = BtOpGetGreaterEqual
        Else
            
'            If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Then
'            Else
            
            
                                                        '標準棚番
                                
                Edit = StrConv(SOKOREC.Soko_No, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
                Write #FileNo, Edit,
                                                        '品番（外部）

                Write #FileNo, StrConv(GOODSREC.HIN_GAI, vbUnicode),
                                                        '箱№
                Write #FileNo, StrConv(GOODSREC.PACKING_NO, vbUnicode),
                                                        '商品化済み在庫数
                Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Write #FileNo, Edit,
                                                        '未商品在庫数
                Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Write #FileNo, Edit,
                                                        
                If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                End If
                                                        '未商品別置き
                If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) = 0 Then
                    Write #FileNo, ,
                Else
                    Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(0).EE_LOC, 2)
                    Edit = Edit & " " & Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                    Write #FileNo, Edit,
                End If
                                                        
                If Len(Trim(EE_ZAIKO_TBL(1).EE_LOC)) = 0 Then
                    Write #FileNo, ,
                Else
                    Edit = Left(EE_ZAIKO_TBL(1).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(1).EE_LOC, 2)
                    Edit = Edit & " " & Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                    Write #FileNo, Edit,
                End If
                                                        
                If Len(Trim(EE_ZAIKO_TBL(2).EE_LOC)) = 0 Then
                    Write #FileNo, ,
                Else
                    Edit = Left(EE_ZAIKO_TBL(2).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(2).EE_LOC, 2)
                    Edit = Edit & " " & Format(EE_ZAIKO_TBL(2).EE_QTY, "#0")
                    Write #FileNo, Edit,
                End If
                                                        
                                                        '月平均出荷数
                Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Write #FileNo, Edit,
                                                        '事前商品化必要数
                Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Write #FileNo, Edit,
                                                        '事前商品化状況
                Edit = Format(CInt(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Write #FileNo, Edit
                
'            End If
            com = BtOpGetNext
        End If
    Loop

    Close #FileNo

    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    Call Input_UnLock
    
    Data_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Proc = True
    End If


End Function

Private Function MI_ZAIKO_KENSAKU(Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   未商品の処理
'----------------------------------------------------------------------------
Dim i           As Integer

Dim com         As Integer
Dim sts         As Integer

    MI_ZAIKO_KENSAKU = True
    
    For i = 0 To UBound(EE_ZAIKO_TBL)
        EE_ZAIKO_TBL(i).EE_LOC = ""
        EE_ZAIKO_TBL(i).EE_QTY = 0
    Next i
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
                    Exit Do
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select
        For i = 0 To UBound(EE_ZAIKO_TBL)
                        
            If Trim(EE_ZAIKO_TBL(i).EE_LOC) = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                Exit For
            Else
                If Len(Trim(EE_ZAIKO_TBL(i).EE_LOC)) = 0 Then
                    EE_ZAIKO_TBL(i).EE_LOC = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                    Exit For
                End If
            End If
        Next i
    
        If i > UBound(EE_ZAIKO_TBL) Then
            Exit Do
        End If
            
    
        EE_ZAIKO_TBL(i).EE_QTY = EE_ZAIKO_TBL(i).EE_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
    
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function

Public Function F106027_Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional LOCATION As String = "        ") As Integer
'****************************************************
'*      在庫数集計
'*
'*  品番または品番＋棚番毎の在庫数を集計する。
'*
'*  引数 :  在庫数（商品化済み）
'*          在庫数（未商品）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番(省略可 省略=空白)
'*  戻り値: false    正常
'*          SYS_ERR  継続できない異常
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Soko_No     As String * 2
Dim Retu        As String * 2
Dim Ren         As String * 2
Dim Dan         As String * 2
    
Dim Not_GOODS   As Boolean

Dim i           As Integer

    F106027_Zaiko_Syukei_Proc = SYS_ERR

    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreater

    If Len(Trim(LOCATION)) = 0 Then
                                '倉庫番号空白は棚番省略とみなす
        Call UniCode_Conv(K1_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K1_ZAIKO.Retu, "")
        Call UniCode_Conv(K1_ZAIKO.Ren, "")
        Call UniCode_Conv(K1_ZAIKO.Dan, "")

        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop

    Else

        Soko_No = Mid(LOCATION, 1, 2)
        Retu = Mid(LOCATION, 3, 2)
        Ren = Mid(LOCATION, 5, 2)
        Dan = Mid(LOCATION, 7, 2)

        Call UniCode_Conv(K0_ZAIKO.Soko_No, Soko_No)
        Call UniCode_Conv(K0_ZAIKO.Retu, Retu)
        Call UniCode_Conv(K0_ZAIKO.Ren, Ren)
        Call UniCode_Conv(K0_ZAIKO.Dan, Dan)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(Retu)) = 0 Then
                        Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
                    End If
                    If Len(Trim(Ren)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
                    End If
                    If Len(Trim(Dan)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Dan, vbUnicode)
                    End If

                    If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop
    End If

    F106027_Zaiko_Syukei_Proc = False

End Function



Private Function Data_Make_Sub(Save_Hin_No As String, Syuka_Qty As Long) As Integer
    
Dim sts         As Integer
Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
Dim AVE_QTY     As Long
Dim ans         As Integer
    
    
    Data_Make_Sub = True
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Save_Hin_No)


    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select

    '-----------------------------------------  '商品化集計ファイル作成
    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                
                                                '事業部
        Call UniCode_Conv(GOODSREC.JGYOBU, Last_JGYOBU)
                                                '国内外
        Call UniCode_Conv(GOODSREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                                                '品番（外部）
        Call UniCode_Conv(GOODSREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                '標準棚番
        Call UniCode_Conv(GOODSREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
        Call UniCode_Conv(GOODSREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
        Call UniCode_Conv(GOODSREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
        Call UniCode_Conv(GOODSREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                '箱№
        Call UniCode_Conv(GOODSREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
        
                                                '在庫集計処理
        If F106027_Zaiko_Syukei_Proc(Sumi_QTY, _
                                Mi_QTY, _
                                Last_JGYOBU, _
                                Right(Combo(pcmbNaigai).Text, 1), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
            Exit Function
        End If
                                                
                                                
        If Sumi_QTY > 0 Then
            Sumi_QTY = Sumi_QTY - 1             'サンプル分マイナス
        End If
                                                '商品化済み在庫数
        Call UniCode_Conv(GOODSREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                                '未商品在庫数
        Call UniCode_Conv(GOODSREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                                '月平均出荷数←当日出荷数をセット
        
        AVE_QTY = Syuka_Qty
        Call UniCode_Conv(GOODSREC.AVE_SYUKA, Format(AVE_QTY, "00000000"))
                                                '事前商品化状況
        If AVE_QTY = 0 Then
            Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")
        Else
            Call UniCode_Conv(GOODSREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
        End If
        
        
        Do
            
            sts = BTRV(BtOpInsert, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<GOODS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "商品化支援集計データ")
                    Exit Function
            End Select
        
        Loop
    End If

    Data_Make_Sub = False


End Function
