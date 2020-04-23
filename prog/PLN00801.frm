VERSION 5.00
Begin VB.Form PLN00801 
   Caption         =   "[PLN0080]商品化予定出庫表発行"
   ClientHeight    =   5685
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   9975
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
   OLEDropMode     =   1  '手動
   ScaleHeight     =   5685
   ScaleWidth      =   9975
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "標準棚番一括設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7680
      TabIndex        =   13
      ToolTipText     =   "印刷処理を実行します"
      Top             =   0
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   3
      Left            =   6000
      TabIndex        =   11
      Top             =   3120
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   2
      Left            =   5160
      TabIndex        =   9
      Top             =   3120
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   2400
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1332
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   0
      ItemData        =   "PLN00801.frx":0000
      Left            =   4200
      List            =   "PLN00801.frx":000A
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印 刷"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "印刷処理を実行します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label lblBC 
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8160
      TabIndex        =   12
      Top             =   5040
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   252
      Index           =   4
      Left            =   5640
      TabIndex        =   10
      Top             =   3240
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "標準棚番"
      Height          =   252
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   252
      Index           =   0
      Left            =   5640
      TabIndex        =   6
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "商品化予定日"
      Height          =   252
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "ＢＵ"
      Height          =   252
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   492
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "印刷"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbBU% = 0


Private Const ptxYOTEI_DT_S% = 0
Private Const ptxYOTEI_DT_E% = 1
Private Const ptxST_SOKO_S% = 2
Private Const ptxST_SOKO_E% = 3



Dim NormalFont As New StdFont               '印刷フォント
Dim Code39Font As New StdFont               '印刷フォント

Private KASO_NYUKA_SOKO As String * 2       '仮想　入荷倉庫番号
Private KASO_SYOHN_SOKO As String * 2       '仮想　商品化倉庫番号
Private KASO_NAI_SOKO As String * 2         '仮想　内職倉庫番号


Private Const LMAX% = 55                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）



Private Tanaban()       As String           '別置棚番                               2012.02.14
Private Tana_Sort       As Integer          '別置棚番出力順                         2012.02.14
Private Tana_SU         As Integer          '別置き棚番出力件数                     2012.02.14
Private Ex_SOKO         As Variant          '商品化済みを未商品とみなす仮想倉庫     2012.02.14


Private NYUKA_SOKO_NO   As String           '入荷倉庫№                             2012.02.14

Private Const LAST_UPDATE_DAY$ = " 2012.09.29 11:00"


Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '印刷処理


            If Trim(Text1(ptxYOTEI_DT_S).Text) = "" Then
                Text1(ptxYOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxYOTEI_DT_S).Text) Then
                MsgBox "入力した項目はエラーです。再入力して下さい（商品化予定日　開始）"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If


            If Trim(Text1(ptxYOTEI_DT_E).Text) = "" Then
                Text1(ptxYOTEI_DT_E).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxYOTEI_DT_E).Text) Then
                MsgBox "入力した項目はエラーです。再入力して下さい（商品化予定日　終了）"
                Text1(ptxYOTEI_DT_E).SetFocus
                Exit Sub
            End If


            If Text1(ptxYOTEI_DT_S).Text > Text1(ptxYOTEI_DT_E).Text Then
                MsgBox "入力した項目はエラーです。再入力して下さい（商品化予定日）"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If


            If Trim(Text1(ptxST_SOKO_E).Text) = "" Then
                Text1(ptxST_SOKO_E).Text = "zz"
            End If

            If Text1(ptxST_SOKO_S).Text > Text1(ptxST_SOKO_E).Text Then
                MsgBox "入力した項目はエラーです。再入力して下さい（標準棚番）"
                Text1(ptxST_SOKO_S).SetFocus
                Exit Sub
            End If


            If Print_Proc() Then
                Unload Me
            End If

        Case 1          '終了

            Unload Me
    
    
        Case 2          '標準棚番一括設定                   2012.09.29
    
            If ST_TANABAN_PROC() Then   '標準棚番再設定     2012.09.29
                Unload Me
            End If
    
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]商品化リスト印刷", Me.hwnd, 0)
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
    LOG_F = RTrim(c)

                                
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                




    PLN00801.Caption = PLN00801.Caption & " " & LAST_UPDATE_DAY


    Call Bu_Set_Proc
                                
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.02.14 INI関係追加
                                '入荷仮想倉庫取り込み
'    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", App.EXEName, c) Then
'        c = ""
'    End If
'    KASO_NYUKA_SOKO = RTrim(c)
                                '商品化仮想倉庫取り込み
'    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", App.EXEName, c) Then
'        c = ""
'    End If
'    KASO_SYOHN_SOKO = RTrim(c)
                                '内職仮想倉庫取り込み
'    If GetIni(App.EXEName, "KASO_NAI_SOKO", App.EXEName, c) Then
'        c = ""
'    End If
'    KASO_NAI_SOKO = RTrim(c)

                                '棚番出力順
    If GetIni(App.EXEName, "TANA_SORT", App.EXEName, c) Then
        c = "0"
    End If
    If IsNumeric(c) Then
        Tana_Sort = 0
    Else
        Tana_Sort = Val(c)
        If Tana_Sort < 0 Or Tana_Sort > 2 Then
            Tana_Sort = 0
        End If
    End If
                                '別置き棚番出力件数
    If GetIni(App.EXEName, "TANA_SU", App.EXEName, c) Then
        c = "1"
    End If
    If Not IsNumeric(c) Then
        Tana_SU = 1
    Else
        Tana_SU = Val(c)
        If Tana_SU <= 0 Then
            Tana_SU = 1
        End If
    End If
                                '商品化済みを未商品とみなす仮想倉庫
    If GetIni(App.EXEName, "Ex_SOKO", App.EXEName, c) Then
        c = "**"
    End If
    Ex_SOKO = Split(Trim(c), ",", -1)

                                '通常入荷の要因
    If GetIni(App.EXEName, "NYUKA_SOKO_NO", App.EXEName, c) Then
        c = "**"
    Else
        NYUKA_SOKO_NO = Trim(c)
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.02.14 INI関係追加
                                
                                '商品化予定ファイルＯＰＥＮ
    If PLN_S_YOTEI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '商品化指図データ（親）ＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenRead) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ                 2012.02.14
    If IDO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '商品化用入荷予定ファイルＯＰＥＮ   2012.02.14
    If PLN_Y_NYUKA_Open(BtOpenRead) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ                 2012.09.29
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If




                                '印刷フォント設定
    With NormalFont
        .NAME = PLN00801.FontName
        .Size = 10
    End With
                                '印刷フォント設定（バーコード）
    With Code39Font
        .NAME = lblBC.FontName
        .Size = lblBC.FontSize
    End With

    Text1(ptxYOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxYOTEI_DT_E).Text = Format(Now, "YYYY/MM/DD")




End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化予定ファイル")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PLN00801 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
    
    End Select



End Sub


Private Function xPrint_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化リスト」明細印刷処理
'
'           2012.02.14  ﾚｲｱｳﾄ変更により廃止
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    

    
Dim Skip_FLg        As Boolean
    
    
Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 2
Dim Betu_LOCATION   As String * 8


Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long
Dim RetBuf          As String
    
Dim i               As Integer

Dim c               As String * 128
Dim SHIMUKE_CODE    As String * 2

    xPrint_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　処理開始！！", Me.hwnd, 0)
    
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    
    
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_SOKO, Text1(ptxST_SOKO_S).Text)
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_RETU, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_REN, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_DAN, "")
    
    Call UniCode_Conv(K5_PLN_S_YOTEI.JGYOBU, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.NAIGAI, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.HIN_GAI, "")
    
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K5_PLN_S_YOTEI, Len(K5_PLN_S_YOTEI), 5)
    
        Select Case sts
            Case BtNoErr
                
                
                
                If StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) > Text1(ptxST_SOKO_E).Text Then
                    Exit Do
                End If
            
            
                Skip_FLg = False
                If Trim(Right(Combo1(pcmbBU).Text, 1)) <> "" Then
                    If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) <> Right(Combo1(pcmbBU).Text, 1) Then
                        Skip_FLg = True
                    End If
                End If
            
            
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) < Format(Text1(ptxYOTEI_DT_S).Text, "YYYYMMDD") Or _
                    StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(Text1(ptxYOTEI_DT_E).Text, "YYYYMMDD") Then
                    Skip_FLg = True
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "商品化予定データ")
                Exit Function
        End Select
            
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.15 指図票発行日付＆完了（分納）日付の獲得
        If GetIni(App.EXEName, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), App.EXEName, c) Then
            SHIMUKE_CODE = ""
        Else
            SHIMUKE_CODE = Trim(c)
        End If
        
        
        
        
        Call UniCode_Conv(K4_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
        Call UniCode_Conv(K4_P_SSHIJI_O.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.Print_datetime, StrConv(PLN_S_YOTEI_R.Ins_DateTime, vbUnicode))
        
        com = BtOpGetGreaterEqual
        
        Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, "")
        Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, "")
        
        
        Do
            DoEvents
            sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K4_P_SSHIJI_O, Len(K4_P_SSHIJI_O), 4)
            Select Case sts
                Case BtNoErr
                        
                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Then
                        Exit Do
                    End If
                        
                    If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                        
                        
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                    Else
                        Skip_FLg = True
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "商品化指図データ（親）")
                    Exit Function
            End Select
        
            com = BtOpGetNext
        
        Loop
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.15 指図票発行日付＆完了（分納）日付の獲得
        If Skip_FLg Then
        Else
            
            
 '           If Lcnt = 99 Then
 '               SAVE_SOKO_No = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)
 '           Else
 '                                               '倉庫のブレーク
 '               If SAVE_SOKO_No <> StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) Then
 '                   Lcnt = LMAX + 1
 '                   SAVE_SOKO_No = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)
 '               End If
 '           End If
        
        
        
            If Lcnt > LMAX Then                 'ヘッダーコントロール
                If Head_Proc(Lcnt, SAVE_SOKO_No) Then
                    Exit Function
                End If
            End If
        
        
            Printer.Print Tab(MGN_L);
                                            '標準棚番
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode);
                                            'BU
            Printer.Print Tab(MGN_L + 13);
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    Printer.Print Left(JGYOBU_T(i).NAME, 10);
                End If
            Next i
                                            '品番
            Printer.Print Tab(MGN_L + 28);
            Printer.Print Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode));
                                            '標準棚　在庫数
            If Trim(StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)) = "" Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
            Printer.Print Tab(MGN_L + 43);
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '別置棚検索
'>>>>>>>>>>>>>>>            If Tana_Kensaku(Betu_LOCATION) Then
'>>>>>>>>>>>>>>>                xPrint_Proc = True
'>>>>>>>>>>>>>>>                Exit Function
'>>>>>>>>>>>>>>>            End If
                                            '別置棚　在庫数
            If Trim(Betu_LOCATION) <> "" Then
                Printer.Print Tab(MGN_L + 54);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
            
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    Betu_LOCATION) Then
                    Exit Function
                End If
                Printer.Print Tab(MGN_L + 65);
                ZAIKO_QTY = SUMI_QTY + MI_QTY
                RetBuf = Format(ZAIKO_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
            End If
                                            '商品化室＆内職　在庫数
            Printer.Print Tab(MGN_L + 80);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    KASO_NAI_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '入荷倉庫在庫
            Printer.Print Tab(MGN_L + 90);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '予定数
            Printer.Print Tab(MGN_L + 100);
            RetBuf = Format(Val(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            Printer.Print Tab(MGN_L + 113);
                                
                                '印刷フォント設定（CODE39）
            Set Printer.Font = Code39Font
                                'バーコード(*品番*)
            Printer.Print "*" & Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) & "*";
                                '印刷フォント設定（通常）
            Set Printer.Font = NormalFont
            
            Printer.Print
            Printer.Print
        
        
            Lcnt = Lcnt + 2
        
        End If
            
            
            
    
        com = BtOpGetNext
    
    
    Loop


    Printer.EndDoc


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　処理終了！！", Me.hwnd, 0)





    xPrint_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化リスト」明細印刷処理
'
'           2012.02.14  新ﾚｲｱｳﾄ対応
'
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    

    
Dim Skip_FLg        As Boolean
    
    
Dim Lcnt                As Integer
Dim SAVE_SOKO_No        As String * 2
Dim Betu_LOCATION       As String * 8


Dim SUMI_QTY            As Long
Dim MI_QTY              As Long
Dim ZAIKO_QTY           As Long
Dim TEMP_QTY            As Long


Dim wkNYUKA_YOTEI_QTY   As Long     '2012.02.14

Dim RetBuf              As String
    
Dim i                   As Integer
Dim j                   As Integer  '2012.02.14

Dim FOUND_F             As Boolean  '2012.02.14


Dim c                   As String * 128
Dim SHIMUKE_CODE        As String * 2

    Print_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　処理開始！！", Me.hwnd, 0)
    
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    
    If Trim(Right(Combo1(pcmbBU).Text, 1)) = "" Then
        com = BtOpGetFirst
    
    
    Else
    
        Call UniCode_Conv(K6_PLN_S_YOTEI.JGYOBU, Trim(Right(Combo1(pcmbBU).Text, 1)))
        Call UniCode_Conv(K6_PLN_S_YOTEI.NAIGAI, "1")
    
    
        Call UniCode_Conv(K6_PLN_S_YOTEI.ST_SOKO, Text1(ptxST_SOKO_S).Text)
        Call UniCode_Conv(K6_PLN_S_YOTEI.ST_RETU, "")
        Call UniCode_Conv(K6_PLN_S_YOTEI.ST_REN, "")
        Call UniCode_Conv(K6_PLN_S_YOTEI.ST_DAN, "")
        
        Call UniCode_Conv(K6_PLN_S_YOTEI.HIN_GAI, "")
    
        com = BtOpGetGreaterEqual
    
    End If

    Do
        DoEvents
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K6_PLN_S_YOTEI, Len(K6_PLN_S_YOTEI), 6)
    
        Select Case sts
            Case BtNoErr
                
                Skip_FLg = False
                
                If Trim(Right(Combo1(pcmbBU).Text, 1)) <> "" Then
                    If Trim(StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode)) <> Trim(Right(Combo1(pcmbBU).Text, 1)) Then
                        Exit Do
                    End If
                End If
                
                If Trim(StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode)) <> "" Then
                    If StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) > Text1(ptxST_SOKO_E).Text Then
                        Exit Do
                    End If
                Else
                    If StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) < Text1(ptxST_SOKO_S).Text Or _
                        StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) > Text1(ptxST_SOKO_E).Text Then
                        
                        Skip_FLg = True
                    
                    End If
                End If
            
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) < Format(Text1(ptxYOTEI_DT_S).Text, "YYYYMMDD") Or _
                    StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(Text1(ptxYOTEI_DT_E).Text, "YYYYMMDD") Then
                    Skip_FLg = True
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "商品化予定データ")
                Exit Function
        End Select
            
        If GetIni(App.EXEName, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), App.EXEName, c) Then
            SHIMUKE_CODE = ""
        Else
            SHIMUKE_CODE = Trim(c)
        End If
        
        
        
        
        Call UniCode_Conv(K4_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
        Call UniCode_Conv(K4_P_SSHIJI_O.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.Print_datetime, StrConv(PLN_S_YOTEI_R.Ins_DateTime, vbUnicode))
        
        com = BtOpGetGreaterEqual
        
        Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, "")
        Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, "")
        
        
        Do
            DoEvents
            sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K4_P_SSHIJI_O, Len(K4_P_SSHIJI_O), 4)
            Select Case sts
                Case BtNoErr
                        
                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Then
                        Exit Do
                    End If
                        
                    If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                        
                        
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                    Else
                        Skip_FLg = True
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "商品化指図データ（親）")
                    Exit Function
            End Select
        
            com = BtOpGetNext
        
        Loop
        
        
If Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) = "ARR31-628" Then
    Debug.Print
End If
        
        
        If Skip_FLg Then
        Else
        
            If Lcnt > LMAX Then                 'ヘッダーコントロール
                If Head_Proc(Lcnt, SAVE_SOKO_No) Then
                    Exit Function
                End If
            End If
        
        
            Printer.Print Tab(MGN_L);
                                            'BU
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    Printer.Print Left(JGYOBU_T(i).NAME, 10);
                End If
            Next i
            Printer.Print Tab(MGN_L + 13);
                                            '標準棚番
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode);
                                            '品番
            Printer.Print Tab(MGN_L + 28);
            Printer.Print Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode));
                                            '標準棚　在庫数
            If Trim(StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)) = "" Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
            Printer.Print Tab(MGN_L + 43);
            RetBuf = Format(SUMI_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            RetBuf = Format(MI_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '別置棚検索
            Erase Tanaban
            ReDim Tanaban(0 To 0)
            Tanaban(0) = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & _
                            StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & _
                            StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & _
                            StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode)
            
            If Tana_Kensaku() Then
                Print_Proc = True
                Exit Function
            End If
                                            '別置棚　在庫数
            If UBound(Tanaban) > 0 Then
                Printer.Print Tab(MGN_L + 66);
                Printer.Print Left(Tanaban(1), 2) & "-" _
                                & Mid(Tanaban(1), 3, 2) & "-" _
                                & Mid(Tanaban(1), 5, 2) & "-" _
                                & Right(Tanaban(1), 2);
            
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    Tanaban(1)) Then
                    Exit Function
                End If
                Printer.Print Tab(MGN_L + 80);
                
                For j = 0 To UBound(Ex_SOKO)
                    If Left(Tanaban(1), 2) = Ex_SOKO(j) Then
                        MI_QTY = MI_QTY + SUMI_QTY
                        SUMI_QTY = 0
                        Exit For
                    End If
                Next j
                
                
                
                RetBuf = Format(SUMI_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                RetBuf = Format(MI_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
            End If
                                            '予定数
            Printer.Print Tab(MGN_L + 105);
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  入荷予定数の計算 2012.02.14
'            RetBuf = Format(Val(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,##0")
            Call UniCode_Conv(K1_PLN_Y_NYUKA.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
            
'''''''''''''''''''''   2012.03.15
'            Call UniCode_Conv(K1_PLN_Y_NYUKA.HIN_GAI, StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
'''''''''''''''''''''   2012.03.15
            
            Call UniCode_Conv(K1_PLN_Y_NYUKA.N_YOTEI_DT, Format(Now, "YYYYMMDD"))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.SEQ_NO, "")
            
            com = BtOpGetGreaterEqual
            
            wkNYUKA_YOTEI_QTY = 0
            
            Do
                DoEvents
                sts = BTRV(com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K1_PLN_Y_NYUKA, Len(K1_PLN_Y_NYUKA), 1)
                Select Case sts
                    Case BtNoErr
                            
                            
                        If StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                            StrConv(PLN_Y_NYUKA_R.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                            Trim(StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                            Exit Do
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "商品化用入荷予定ファイル")
                        Exit Function
                End Select
            
            
                If Trim(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode)) = "" Then
                    wkNYUKA_YOTEI_QTY = wkNYUKA_YOTEI_QTY + Val(StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode))
                Else
                    If Trim(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode)) < Format(Now, "YYYYMMDD") Then
                    Else
                        If Trim(StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode)) = Format(Now, "YYYYMMDD") Then
                            Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode))
                            Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(PLN_Y_NYUKA_R.NAIGAI, vbUnicode))
                            Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode))
                            Call UniCode_Conv(K1_IDO.JITU_DT, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(K1_IDO.JITU_TM, "")
                            com = BtOpGetGreaterEqual
    
                            FOUND_F = False
                            Do
                                DoEvents
                                sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                                Select Case sts
                                    Case BtNoErr
                                        If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) Or _
                                            StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(PLN_Y_NYUKA_R.NAIGAI, vbUnicode) Or _
                                            StrConv(IDOREC.HIN_GAI, vbUnicode) <> StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode) Or _
                                            StrConv(IDOREC.JITU_DT, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                                            
                                            Exit Do
                                
                                        End If
                                        If NYUKA_SOKO_NO = StrConv(IDOREC.TO_SOKO, vbUnicode) Then
                                            FOUND_F = True
                                            Exit Do
                                        End If
                                End Select
                            
                                com = BtOpGetNext
                            
                            Loop
                            
                            If Not FOUND_F Then
                                wkNYUKA_YOTEI_QTY = wkNYUKA_YOTEI_QTY + Val(StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode))
                            End If
                        Else
                        
                            wkNYUKA_YOTEI_QTY = wkNYUKA_YOTEI_QTY + Val(StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode))
                        End If
                    End If
                End If
            
            
                com = BtOpGetNext
                    
            Loop
            
            RetBuf = Format(wkNYUKA_YOTEI_QTY, "#,##0")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  入荷予定数の計算 2012.02.14
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            
            
            Printer.Print Tab(MGN_L + 117);
            RetBuf = Format(Val(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            
            
            Printer.Print Tab(MGN_L + 128);
                                
                                '印刷フォント設定（CODE39）
            Set Printer.Font = Code39Font
                                'バーコード(*品番*)
            Printer.Print "*" & Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) & "*";
                                '印刷フォント設定（通常）
            Set Printer.Font = NormalFont
            Printer.Print
            Lcnt = Lcnt + 1
'---------------------------------------------------------------------  別置き棚番  2012.02.14
            For i = 2 To Tana_SU
            
                If i > UBound(Tanaban) Then
                    Exit For
                End If
            
            
                If Lcnt > LMAX Then                 'ヘッダーコントロール
                    If Head_Proc(Lcnt, SAVE_SOKO_No) Then
                        Exit Function
                    End If
                End If
            
            
                Printer.Print Tab(MGN_L + 66);
                Printer.Print Left(Tanaban(i), 2) & "-" _
                                & Mid(Tanaban(i), 3, 2) & "-" _
                                & Mid(Tanaban(i), 5, 2) & "-" _
                                & Right(Tanaban(i), 2);
            
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    Tanaban(i)) Then
                    Exit Function
                End If
                
                For j = 0 To UBound(Ex_SOKO)
                    If Left(Tanaban(i), 2) = Ex_SOKO(j) Then
                        MI_QTY = MI_QTY + SUMI_QTY
                        SUMI_QTY = 0
                        Exit For
                    End If
                Next j
                
                
                Printer.Print Tab(MGN_L + 80);
                RetBuf = Format(SUMI_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                RetBuf = Format(MI_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf
                Lcnt = Lcnt + 1
            
            Next i
'---------------------------------------------------------------------  別置き棚番  2012.02.14
            Printer.Print
            Lcnt = Lcnt + 1
        
        End If
            
            
            
    
        com = BtOpGetNext
    
    
    Loop


    Printer.EndDoc


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　処理終了！！", Me.hwnd, 0)





    Print_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function



Private Function Head_Proc(Lcnt As Integer, SAVE_SOKO_No As String) As Integer
'----------------------------------------------------------------------------
'                   「商品化リスト」ヘッダー印刷処理
'----------------------------------------------------------------------------
Dim i As Integer
Dim sts As Integer

    Head_Proc = True

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    'Printer.Print Tab(3);                  '97.10.14
    
    Printer.Print Tab(MGN_L + 41);
    
    Printer.Print "『商品化予定出庫表』     　　商品化予定日：";
    Printer.Print Text1(ptxYOTEI_DT_S).Text & "～" & Text1(ptxYOTEI_DT_E).Text;
    
    Printer.Print Tab(MGN_L + 110);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

'    Printer.Print Tab(MGN_L + 5);
'    Printer.Print "倉庫：";
'    Printer.Print SAVE_SOKO_No;
'    Printer.Print Tab(MGN_L + 15);
'    Call UniCode_Conv(K0_SOKO.Soko_No, SAVE_SOKO_No)
'    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'    Select Case sts
'        Case BtNoErr
'            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
'        Case BtErrKeyNotFound
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'            Exit Function
'    End Select
'
'    Printer.Print

    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.02.14
'    Printer.Print Tab(MGN_L);
'    Printer.Print "標準棚番";
'    Printer.Print Tab(MGN_L + 13);
'    Printer.Print "ＢＵ";
'    Printer.Print Tab(MGN_L + 28);
'    Printer.Print "品番（外部）";
'    Printer.Print Tab(MGN_L + 42);
'    Printer.Print "標準棚在庫";
'    Printer.Print Tab(MGN_L + 54);
'    Printer.Print "別置棚番";
'    Printer.Print Tab(MGN_L + 81);
'    Printer.Print "商品化室";
'    Printer.Print Tab(MGN_L + 91);
'    Printer.Print "入荷倉庫";
'    Printer.Print Tab(MGN_L + 101);
'    Printer.Print "　予定数";
'    Printer.Print


    Printer.Print Tab(MGN_L);
    Printer.Print "ＢＵ";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 46);
    Printer.Print "標準棚在庫数";
    Printer.Print Tab(MGN_L + 66);
    Printer.Print "別置棚番";
    Printer.Print Tab(MGN_L + 81);
    Printer.Print "別置棚番在庫数";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "入荷予定数";
    Printer.Print Tab(MGN_L + 114);
    Printer.Print "商品化予定数"
    
    
    
    Printer.Print Tab(MGN_L + 48);
    Printer.Print "商";
    Printer.Print Tab(MGN_L + 56);
    Printer.Print "未商";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "商";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "未商"
    Printer.Print

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.02.14

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku() As Integer
'----------------------------------------------------------------------------
'                   別置き棚番検索
'
'           検索順を入荷日順から棚番順に変更(KEY6-->KEY1)    2012.02.14
'----------------------------------------------------------------------------

Dim sts     As Integer

Dim i       As Integer
Dim Skip_F  As Boolean


    Tana_Kensaku = True
    
'>>>>>>>>>>>>    Betu_LOCATION = ""
    
    
If Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) = "AUA02-1201-W" Then
    Debug.Print
End If
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_ON)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                    
                    Exit Do
                
                End If
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.14
'                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) Or _
'                   StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) Or _
'                   StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) Or _
'                   StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode) Then
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.14
                                                
                                                
                                                
                                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
'                                                'システム倉庫の判定
'                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
'                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Dan, vbUnicode)
'                                Exit Do
'
'                            End If
'                        Case BtErrKeyNotFound
'                                                '考えられないので読み飛ばし
'                        Case Else
'                            Call File_Error(sts, BtOpGetGreater, "倉庫マスタ")
'                            Exit Function
'                    End Select
'
'                                                'システム倉庫を有効とする
'                    If KASO_NYUKA_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
'                        KASO_SYOHN_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
'                        KASO_NAI_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
'                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Dan, vbUnicode)
'                                Exit Do
'
'                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.14
'                End If
                
                Skip_F = True
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                    For i = 0 To UBound(Ex_SOKO)
                        If Ex_SOKO(i) = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                                Skip_F = False
                                Exit For
                            End If
                        End If
                    Next i
                Else
                    Skip_F = False
                End If
                
                If Not Skip_F Then
                    For i = 0 To UBound(Tanaban)
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) = _
                            Tanaban(i) Then
                            Exit For
                        End If
                    Next i
                    If i > UBound(Tanaban) Then
                        ReDim Preserve Tanaban(0 To i)
                        Tanaban(i) = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
'>>>>>>>>>>>>                        Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                    End If
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.02.14
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
            
    Loop
    
    Tana_Kensaku = False

End Function



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00801.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00801)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00801)


    PLN00801.MousePointer = vbDefault

End Sub



Private Sub Bu_Set_Proc()
'----------------------------------------------------------------------------
'                   画面項目（ＢＵ）のセット
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbBU).Clear


        Combo1(pcmbBU).AddItem "全て" & "          " & " "
        



    For i = 0 To UBound(JGYOBU_T)
            
        Combo1(pcmbBU).AddItem JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
            
            
    Next i

    Combo1(pcmbBU).ListIndex = 0
End Sub



Private Function ST_TANABAN_PROC() As Integer
'----------------------------------------------------------------------------
'                   標準棚番のセット
'----------------------------------------------------------------------------
Dim com As Integer
Dim sts As Integer


    ST_TANABAN_PROC = True
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　標準棚番再設定中！！", Me.hwnd, 0)



    com = BtOpGetFirst



    Do
        DoEvents
        
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化予定ファイル")
                Call Input_UnLock
                Exit Function
        End Select
    
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
            Case BtErrKeyNotFound
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_SOKO, "")
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_RETU, "")
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_REN, "")
                Call UniCode_Conv(PLN_S_YOTEI_R.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                Call Input_UnLock
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpUpdate, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "商品化予定ファイル")
                Call Input_UnLock
                Exit Function
        End Select
    
    
        com = BtOpGetNext
    Loop

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "商品化予定出庫表発行　標準棚番再設定終了！！", Me.hwnd, 0)


    Call Input_UnLock
    ST_TANABAN_PROC = False
End Function
