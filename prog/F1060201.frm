VERSION 5.00
Begin VB.Form F1060201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "商品化計画支援アラームリスト印刷"
   ClientHeight    =   7125
   ClientLeft      =   2325
   ClientTop       =   2715
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   32
      Top             =   2880
      Width           =   1320
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   30
      Top             =   2880
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Index           =   1
      Left            =   7440
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Index           =   0
      Left            =   7440
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INI表示"
      Height          =   372
      Left            =   9600
      TabIndex        =   26
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   22
      Top             =   2400
      Width           =   480
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   20
      Top             =   2400
      Width           =   480
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3120
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1320
      Width           =   375
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
      ToolTipText     =   "出力内容は旧形式のままです"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "ﾃﾞｰﾀ"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "～"
      Height          =   240
      Index           =   7
      Left            =   4440
      TabIndex        =   31
      Top             =   3000
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "入荷予定日"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   29
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "抽出除外条件　OEM品番(出荷区分ZZ) 商品化計画除外ﾌﾗｸﾞ：１"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "抽出対象条件　国内供給区分 1:対象　2:打切案内中 空白"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   6240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "%"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "%　～"
      Height          =   240
      Index           =   2
      Left            =   3600
      TabIndex        =   21
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "事前商品化状況"
      Height          =   252
      Index           =   1
      Left            =   1176
      TabIndex        =   19
      Top             =   2520
      Width           =   1776
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "（※空白：全倉庫指定）"
      Height          =   252
      Left            =   3000
      TabIndex        =   18
      Top             =   1920
      Width           =   2652
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Height          =   252
      Left            =   3600
      TabIndex        =   17
      Top             =   1440
      Width           =   2412
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   252
      Index           =   33
      Left            =   2280
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "標準棚番(倉庫番号)"
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   14
      Top             =   1440
      Width           =   2292
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
      TabIndex        =   13
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
Attribute VB_Name = "F1060201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSOKO% = 0                  '開始　標準棚番　倉庫
Private Const ptxFROM_SUMI_PERCENT% = 1     '事前商品化状況から     2011.07.04
Private Const ptxTO_SUMI_PERCENT% = 2       '事前商品化状況まで     2011.07.04

Private Const ptxN_YOTEI_DT_S% = 3          '入荷予定日　開始       2011.09.15
Private Const ptxN_YOTEI_DT_E% = 4          '入荷予定日　終了       2011.09.15


Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNaigai% = 0               '国内外


Private Const LMAX% = 36                    '頁内最大行数
Private Const LCTL% = 99                    '
Private Const MGN_L% = 3                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private NormalFont  As New StdFont          '印刷フォント
Private MidFont     As New StdFont          '印刷フォント

Private OutSide     As Long                 '印刷対外出荷数

Private GOODS_DATA  As String               '出力データファイル名


Private Type EE_ZAIKO_TBL_tag
    EE_LOC          As String * 8
    EE_QTY          As Long
End Type

Private EE_ZAIKO_TBL(0 To 2) As EE_ZAIKO_TBL_tag


Private SHIMUKE_CODE    As String * 2       '仕向け先コード 2008.03.03


Private SORT_SEQ        As Integer          'SORT順 2008.11.06


'''''''''''''''''''''''''''''''''''''''''''''   2011.03.31
Private Type KOUSEI_TBL
    KO_JGYOBU   As String * 1           '事業部
    KO_NAIGAI   As String * 1           '国内外
    KO_SYUBETSU As String * 2           '種別
    KO_HIN_GAI  As String * 20          '品番
    KO_QTY      As Double               '員数
    G_ST_SHITAN As Double               '仕入＠
    G_ST_URITAN As Double               '売上＠
    G_ST_SHIKIN As Double               '仕入金額
    G_ST_URIKIN As Double               '売上金額
    S_KOUSU     As Double               '作業時間
    SEI_SYU_KON As Double               '集合梱包
    G_ST_URIKIN_KUSATU As _
                    Double              '草津専用
End Type




Dim SHIZAI_T        As Variant          '資材対象
Dim DOUKON_T        As Variant          '同梱対象
Dim KAKOU_T         As Variant          '加工対象

Dim KUSATU_F                As Boolean  '対象センター　草津 OR 草津以外


Dim KOSOU_KBN       As String * 2       '個装区分
Dim GAISO_KBN       As String * 2       '外装区分

'''''''''''''''''''''''''''''''''''''''''''''   2011.03.31



'''''''''''''''''''''''''''''''''''''''''''''   2011.07.04
Dim SAMPLE_QTY      As Integer          '見本除外数
Dim NOT_Hin_Name    As Variant          '除外品名
Dim NOT_Hin_Name_F  As Boolean          '除外品名有無
Dim wkNOT_Hin_Name  As String

Dim TUKI1_TITLE     As String           '月平均出荷数ﾀｲﾄﾙ
Dim S_TUKI1_TITLE   As String           '生産計画用月平均出荷数ﾀｲﾄﾙ(1)
Dim S_TUKI2_TITLE   As String           '生産計画用月平均出荷数ﾀｲﾄﾙ(1)
Dim TUKI1           As Integer
Dim TUKI2           As Integer
Dim TUKI3           As Integer
'''''''''''''''''''''''''''''''''''''''''''''   2011.07.04

Dim GOODS_OFF_SOKO_NO   As Variant      '2011.12.16


Dim CANCEL_F        As Integer          '2012.05.08



'Private Const LAST_UPDATE_DAY$ = "[F106020] 2012.09.24 12:00 [商品化予定対応]"
Private Const LAST_UPDATE_DAY$ = "[F106020] 2018.08.06 18:45 [商品化予定対応]"



Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   エラーチェック処理
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
    If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
        Label2.Caption = "全倉庫"
    Else
        Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Label2.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                        
                If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                    Beep
                    MsgBox "入力した項目はエラーです。（仮想倉庫）"
                    Text(ptxSOKO).SetFocus
                    Err_Chk = False
                    Exit Function
                End If
                    
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（倉庫未登録）"
                Text(ptxSOKO).SetFocus
                Err_Chk = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetGreaterEqual, "倉庫マスタ")
                Exit Function
        End Select
            
            
    End If
    
    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        Text(ptxFROM_SUMI_PERCENT).Text = "000"
    Else
        If Not IsNumeric(Text(ptxFROM_SUMI_PERCENT).Text) Then
            MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
            Text(ptxFROM_SUMI_PERCENT).SetFocus
            Err_Chk = False
            Exit Function
        End If
    End If
    
    '2011.07.04
    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        Text(ptxTO_SUMI_PERCENT).Text = "999"
    Else
        If Not IsNumeric(Text(ptxTO_SUMI_PERCENT).Text) Then
            MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
            Text(ptxTO_SUMI_PERCENT).SetFocus
            Err_Chk = False
            Exit Function
        End If
    End If
    
    If Val(Text(ptxFROM_SUMI_PERCENT).Text) > Val(Text(ptxTO_SUMI_PERCENT).Text) Then
        MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
        Text(ptxFROM_SUMI_PERCENT).SetFocus
        Err_Chk = False
        Exit Function
    End If
    
    '2011.07.04
    
    
    
    
    '2011.09.15
    If Trim(Text(ptxN_YOTEI_DT_S).Text) = "" Then
        Text(ptxN_YOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
    End If
    If Not IsDate(Text(ptxN_YOTEI_DT_S).Text) Then
        MsgBox "入力した項目はエラーです。（商品化予定日（開始））"
        Text(ptxN_YOTEI_DT_S).SetFocus
        Err_Chk = False
        Exit Function
    End If
    
    
    If Trim(Text(ptxN_YOTEI_DT_E).Text) = "" Then
        Text(ptxN_YOTEI_DT_E).Text = "9999/12/31"
    End If
    If Not IsDate(Text(ptxN_YOTEI_DT_E).Text) Then
        MsgBox "入力した項目はエラーです。（商品化予定日（終了））"
        Text(ptxN_YOTEI_DT_E).SetFocus
        Err_Chk = False
        Exit Function
    End If
    
    If Text(ptxN_YOTEI_DT_S).Text > Text(ptxN_YOTEI_DT_E).Text Then
        MsgBox "入力した項目はエラーです。（商品化予定日）"
        Text(ptxN_YOTEI_DT_S).SetFocus
        Err_Chk = False
        Exit Function
    End If
    '2011.09.15
        
    
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060201.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060201)


    F1060201.MousePointer = vbDefault

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
Dim c       As String * 128     '2008.03.03
    
    Select Case Index
        
        Case 7                              'データ出力
            
            
            '2012.05.08
            If CANCEL_F = 0 Then
                CANCEL_F = 1
                Exit Sub
            End If
            '2012.05.08
            
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
                                        '仕向け先取り込み   2008.03.03
            If Last_JGYOBU = "*" Then
            Else
                If GetIni(App.EXEName, Last_JGYOBU, App.EXEName, c) Then
                    MsgBox "仕向け先の設定を行ってください。"
                    Exit Sub
                Else
                    SHIMUKE_CODE = Trim(c)
                End If
            End If
            
            
            Beep
            
            CANCEL_F = 9            '2012.05.08
            
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                
                CANCEL_F = 0        '2012.05.08
                
                
                
                Text1(0).Text = Now
                
                
                If Data_Proc() Then
                    Unload Me
                End If
            
            
            
            
            End If
            
            CANCEL_F = 9            '2012.05.08
            
            
            
            Text(ptxSOKO).SetFocus
        
        
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
                                        '仕向け先取り込み   2008.03.03
            If Last_JGYOBU = "*" Then
                MsgBox "印刷時、全事業部指定は行えません。"
                Exit Sub
            End If
            
            
            If GetIni(App.EXEName, Last_JGYOBU, App.EXEName, c) Then
                MsgBox "仕向け先の設定を行ってください。"
                Exit Sub
            Else
                SHIMUKE_CODE = Trim(c)
            End If
            
            Beep
            yn = MsgBox("「旧印刷内容で印刷されます。」印刷しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxSOKO).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()

    MsgBox "SORT=" & SORT_SEQ & Chr(13) & Chr(10) & _
            "OUTSIDE=" & OutSide & Chr(13) & Chr(10) & _
            "Sample_QTY=" & SAMPLE_QTY & Chr(13) & Chr(10) & _
            "NOT_Hin_Name=" & wkNOT_Hin_Name & Chr(13) & Chr(10)


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
Dim sts As Integer              '2011.03.31
Dim com As Integer              '2011.03.31
     
     
     If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    
    
    
    'ステータスウィンドウを作成する     2011.01.12
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "商品化計画支援アラームリスト", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    
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

    '2011.09.15 全事業部指定可
    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "全事業部"
    '2011.09.15 全事業部指定可
    


    


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
'2011.09.15            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1060201.Caption = "商品化計画支援アラームリスト印刷（" + RTrim(JGYOBU_T(i).NAME) + "）" & LAST_UPDATE_DAY

            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)           '2011.01.12
                                
                                
                                '商品化支援ファイル名取り込み
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'商品化支援ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    GOODS_DATA = Trim(c)
                                
                                
'------------------------------ SYS.INI--> F106020.INI 2011.07.04
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
                                'SORT順取り込み 2008.11.06
    If GetIni(App.EXEName, "SORT", App.EXEName, c) Then
        SORT_SEQ = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            SORT_SEQ = 0
        Else
            SORT_SEQ = CInt(Trim(c))
        End If
    End If
                                
                                '見本数 2011.07.04
    If GetIni(App.EXEName, "Sample_QTY", App.EXEName, c) Then
        SAMPLE_QTY = 0
    Else
        If IsNumeric(Trim(c)) Then
            SAMPLE_QTY = CLng(Trim(c))
        Else
            SAMPLE_QTY = 0
        End If
    End If
                                
                                '品名による除外 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        wkNOT_Hin_Name = Trim(c)
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                
                                
                                
'------------------------------ SYS.INI--> F106020.INI 2011.07.04
                                
                                
'------------------------------------   2011.07.04  平均期間の獲得
    If GetIni(App.EXEName, "TUKI1", "F120050", c) Then
        TUKI1 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI1 = 3
        Else
            TUKI1 = Val(RTrim(c))
        End If
    End If
    TUKI1_TITLE = "月平均出荷数(" & Format(TUKI1, "#0") & "ヵ月)"


    If GetIni(App.EXEName, "TUKI2", "F120050", c) Then
        TUKI2 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI2 = 3
        Else
            TUKI2 = Val(RTrim(c))
        End If
    End If
    S_TUKI1_TITLE = "生産計画用月平均出荷数(" & Format(TUKI2, "#0") & "ヵ月)"


    If GetIni(App.EXEName, "TUKI3", "F120050", c) Then
        TUKI3 = 12
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI3 = 12
        Else
            TUKI3 = Val(RTrim(c))
        End If
    End If
    S_TUKI2_TITLE = "生産計画用月平均出荷数(" & Format(TUKI3, "#0") & "ヵ月)"







'------------------------------------   2011.07.01
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
'-----------------------------------------------    2011.03.31
                                '資材対象種別   2011.03.31
    If GetIni("SEI0010", "SHIZAI", "SEI0010", c) Then
        
        c = "**"
        SHIZAI_T = Split(Trim(c), ",", -1)
        
    Else
        SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                
                                '同梱対象種別
    If GetIni("SEI0010", "DOUKON", "SEI0010", c) Then
        c = "**"
        DOUKON_T = Split(Trim(c), ",", -1)
    Else
        DOUKON_T = Split(Trim(c), ",", -1)
    End If
                                '加工対象種別
   If GetIni("SEI0010", "KAKOU", "SEI0010", c) Then
        c = "**"
        KAKOU_T = Split(Trim(c), ",", -1)
    Else
        KAKOU_T = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                
                                'センターの識別 2011.03.31
    If GetIni("SEI0010", "KUSATU", "SEI0010", c) Then
        KUSATU_F = False
    Else
        If Trim(c) = "1" Then
            KUSATU_F = True
        Else
            KUSATU_F = False
        End If
    End If
                                
                                
                                
                                '個装資材区分の獲得
    If GetIni("SEI0010", "KOSOU", "SEI0010", c) Then
        KOSOU_KBN = ""
    Else
        KOSOU_KBN = Trim(c)
    End If
                                '外装資材区分の獲得
    If GetIni("SEI0010", "GAISO", "SEI0010", c) Then
        GAISO_KBN = ""
    Else
        GAISO_KBN = Trim(c)
    End If
                                
'-----------------------------------------------    2011.03.31
                                
                                
                                
                                '商品化済みで有っても未商品とみなす 2011.12.16
    If GetIni(App.EXEName, "GOODS_OFF_SOKO_NO", App.EXEName, c) Then
        c = "**"
        GOODS_OFF_SOKO_NO = Split(Trim(c), ",", -1)
    Else
        GOODS_OFF_SOKO_NO = Split(Trim(c), ",", -1)
    End If
                                
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                
                                
                                '商品化指図票データＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '受払先マスタＯＰＥＮ   2011.07.04
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                '商品化集計ファイルＯＰＥＮ
    If GOODS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=01)")
        Unload Me
    End Select

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=02)")
        Unload Me
    End Select
                                
                                
                                
                                
                                
                                '商品化用入荷予定ファイルＯＰＥＮ   2011.09.15
    If PLN_Y_NYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                'PNマスタファイルＯＰＥＮ   2012.08.27
    If PN_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                '印刷フォント設定(通常)
    With NormalFont
        .NAME = F1060201.FontName
        .Size = 12
    End With

                                '印刷フォント設定（小）
    With MidFont
        .NAME = F1060201.FontName
        .Size = 8
    End With


    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    CANCEL_F = 9            '2012.05.08

    Show
    
    Text(ptxSOKO).SetFocus
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
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
    
    Set F1060201 = Nothing

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
    F1060201.Caption = "商品化計画支援アラームリスト印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）" & LAST_UPDATE_DAY
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
        Case ptxSOKO
            
            Text(Index).Text = StrConv(Trim(Text(Index).Text), vbUpperCase)
            
            If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Label2.Caption = "全倉庫"
            Else
                Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Label2.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                            Beep
                            MsgBox "入力した項目はエラーです。（仮想倉庫）"
                            Text(ptxSOKO).SetFocus
                            Exit Sub
                        End If
                    
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "入力した項目はエラーです。（倉庫未登録）"
                        Text(ptxSOKO).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetGreaterEqual, "倉庫マスタ")
                        Exit Sub
                End Select
            
            
            End If
    
''''''''''''''''''''''''''''''''''''''''    2011.07.04
'        Case ptxSUMI_PERCENT    '2008.03.03
'
'            If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
'            Else
'                If Not IsNumeric(Text(ptxSUMI_PERCENT).Text) Then
'                    Beep
'                    MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
'                    Text(ptxSUMI_PERCENT).SetFocus
'                    Exit Sub
'                End If
'            End If
    
    
        Case ptxFROM_SUMI_PERCENT

            If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
                Text(ptxFROM_SUMI_PERCENT).Text = "000"
            Else
                If Not IsNumeric(Text(ptxFROM_SUMI_PERCENT).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
                    Text(ptxFROM_SUMI_PERCENT).SetFocus
                    Exit Sub
                End If
            End If
    
        Case ptxTO_SUMI_PERCENT

            If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
                Text(ptxTO_SUMI_PERCENT).Text = "000"
            Else
                If Not IsNumeric(Text(ptxTO_SUMI_PERCENT).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
                    Text(ptxTO_SUMI_PERCENT).SetFocus
                    Exit Sub
                End If
            End If
    
                
            If Val(Text(ptxFROM_SUMI_PERCENT).Text) > Val(Text(ptxTO_SUMI_PERCENT).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。（事前商品化状況（％））"
                Text(ptxFROM_SUMI_PERCENT).SetFocus
                Exit Sub
            End If
    
''''''''''''''''''''''''''''''''''''''''    2011.07.04
    
    
    
    
    
    
''''''''''''''''''''''''''''''''''''''''    2011.09.15
        Case ptxN_YOTEI_DT_S
            If Trim(Text(ptxN_YOTEI_DT_S).Text) = "" Then
                Text(ptxN_YOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text(ptxN_YOTEI_DT_S).Text) Then
                MsgBox "入力した項目はエラーです。（商品化予定日（開始））"
                Text(ptxN_YOTEI_DT_S).SetFocus
                Exit Sub
            End If
        Case ptxN_YOTEI_DT_E

            If Trim(Text(ptxN_YOTEI_DT_E).Text) = "" Then
                Text(ptxN_YOTEI_DT_E).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text(ptxN_YOTEI_DT_E).Text) Then
                MsgBox "入力した項目はエラーです。（商品化予定日（終了））"
                Text(ptxN_YOTEI_DT_E).SetFocus
                Exit Sub
            End If
            
            If Text(ptxN_YOTEI_DT_S).Text > Text(ptxN_YOTEI_DT_E).Text Then
                MsgBox "入力した項目はエラーです。（商品化予定日）"
                Text(ptxN_YOTEI_DT_S).SetFocus
                Exit Sub
            End If

''''''''''''''''''''''''''''''''''''''''    2011.09.15
    
    
    
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
Dim Lcnt                As Integer

Dim sts                 As Integer
Dim com                 As Integer

Dim Save_Soko           As String * 2

Dim Edit                As String

Dim X_Tab               As Integer

'Dim wkSUMI_PERCENT      As Long
Dim wkFROM_SUMI_PERCENT As Long
Dim wkTO_SUMI_PERCENT   As Long


Dim SKIP_F              As Boolean
    
    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    End If


''2011.07.04    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
''2011.07.04        wkSUMI_PERCENT = 100
''2011.07.04    Else
''2011.07.04        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
''2011.07.04    End If

    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        wkFROM_SUMI_PERCENT = 0
    Else
        wkFROM_SUMI_PERCENT = CLng(Text(ptxFROM_SUMI_PERCENT).Text)
    End If


    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        wkTO_SUMI_PERCENT = 999
    Else
        wkTO_SUMI_PERCENT = CLng(Text(ptxTO_SUMI_PERCENT).Text)
    End If

    
        
    
    Lcnt = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    If SORT_SEQ = 0 Then    '2008.11.06

        Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K0_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    Else
    
    
    
        Call UniCode_Conv(K3_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K3_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K3_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K3_GOODS.AVE_SYUKA, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.Sumi_QTY, "")
        Call UniCode_Conv(K3_GOODS.Mi_QTY, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K3_GOODS.HIN_GAI, "")
    End If
    
    
    com = BtOpGetGreater
    
    Do
        
        If SORT_SEQ = 0 Then    '2008.11.06
        
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Else
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K3_GOODS, Len(K3_GOODS), 3)
        End If
        
        Select Case sts
            Case BtNoErr
                
                
                
                
                
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Else
                    If StrConv(GOODSREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                        Exit Do
                    End If
                End If
            
                SKIP_F = False
                If Not IsNumeric(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
''2011.07.04                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
''2011.07.04                        SKIP_F = True
''2011.07.04                    End If
                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) < wkFROM_SUMI_PERCENT Or _
                        CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkTO_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
                
                
                End If
                

                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
                    SKIP_F = True
                End If

                
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select



'-------------------------------------------------  明細印刷
        
        If Not SKIP_F Then
        
        
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
                
            End If
            
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
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
            
            
            If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            '設定発注点より大きい
                Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
                Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzzzzzzzzz")
                com = BtOpGetGreater
            Else
                '未商品在庫＝０ は、印刷対象外 2004.08.27
                
                
                
                
                If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Or _
                    CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
                Else
                    If Head_Print_Proc(Lcnt) Then
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
    
                    Printer.Print StrConv(GOODSREC.HIN_GAI, vbUnicode);
    '                X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 5
                    X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 4
                                                            '箱№
                    Printer.Print Tab(X_Tab);
'2008.11.06                    Printer.Print StrConv(GOODSREC.PACKING_NO, vbUnicode);
    '                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 5
'2008.11.06                    X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                                                            
                                                            
                                                            
                                                            
                    '2008.11.06
                    Printer.Print Left(StrConv(GOODSREC.KOSOU, vbUnicode), 4);
                    X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                    '2008.11.06
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
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
    
                    If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode), StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
    
                    If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
                        Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
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
                
                    Lcnt = Lcnt + 2
            
                End If
            End If
            
            com = BtOpGetNext
        
        End If
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(Lcnt As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If Lcnt < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    If Lcnt = LCTL Then
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

    Printer.Print Tab(MGN_L + 55);
    
    Printer.Print "商品化支援アラームリスト";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫：";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  ";
    Printer.Print "（設定発注点 " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "％）"
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
    Printer.Print Tab(MGN_L + 37);
    Printer.Print "資材";
    Printer.Print Tab(MGN_L + 49);
    Printer.Print "商済数";
    Printer.Print Tab(MGN_L + 61);
    Printer.Print "未商品";
    Printer.Print Tab(MGN_L + 73);
    Printer.Print "月平均";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "必要数";
    Printer.Print Tab(MGN_L + 97);
    Printer.Print "　状況";
    Printer.Print Tab(MGN_L + 120);
    Printer.Print "別置在庫"

    Printer.Print

    Lcnt = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   支援用集計データ作成処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer

Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
Dim AVE_QTY     As Double


'2011.01.13
Dim In_Cnt      As Long
Dim Out_Cnt     As Long
'2011.01.13



'2011.07.04
Dim SKIP_FLG    As Integer
Dim i           As Integer
'2011.07.04


'2011.07.25
Dim Err_Flg     As Integer
Dim c           As String * 128
'2011.07.25



Dim Fast_Flg    As Boolean  '2011.12.16


Dim yn          As Integer  '2012.05.08


    Data_Make_Proc = True

In_Cnt = 0
Out_Cnt = 0


                                            '商品化集計ファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化集計ファイル")
        End If
    End If


    On Error Resume Next
                                            '商品化集計ファイル　フルパス取込み
    sts = GetIni("FILE", GOODS_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [GOODS]読み込みエラー")
        Exit Function
    End If
    Kill (Trim(c))
    On Error GoTo 0
                                '商品化集計ファイルＯＰＥＮ
    If GOODS_Open(BtOpenNomal) Then
        Exit Function
    End If



'---------------------------------------------------------- '全レコード削除
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "集計データ削除中！！", Me.hwnd, 0)
    DoEvents
        
    
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
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
'---------------------------------------------------------- '品目マスタベースでデータ作成

    
    '2011.09.15 全事業部を指定可能とする
    If Last_JGYOBU = "*" Then
        Call UniCode_Conv(K6_ITEM.JGYOBU, "")
        Call UniCode_Conv(K6_ITEM.NAIGAI, "")
        Call UniCode_Conv(K6_ITEM.ST_SOKO, "")
    Else
        Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K6_ITEM.ST_SOKO, Trim(Text(ptxSOKO).Text))
    
    End If
    '2011.09.15 全事業部を指定可能とする
    
    Call UniCode_Conv(K6_ITEM.ST_RETU, "")
    Call UniCode_Conv(K6_ITEM.ST_REN, "")
    Call UniCode_Conv(K6_ITEM.ST_DAN, "")
    Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.08
        DoEvents
        If CANCEL_F = 1 Then
            yn = MsgBox("中断します。よろしいですか？" & Chr(10) & "  はい：処理を中断します。" & Chr(10) & "いいえ：処理を継続します。", vbYesNo, "確認入力")
            If yn = vbYes Then
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "CSVデータ出力中断！！", Me.hwnd, 0)
                Text1(1).Text = Now
                Call Input_UnLock
                Command(7).Enabled = True
                Command(7).Caption = "ﾃﾞｰﾀ"
                CANCEL_F = 9
                Data_Make_Proc = False
                Exit Function
            Else
                CANCEL_F = 0
            End If
        End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.08
        
        
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
        Select Case sts
            Case BtNoErr
                
                
    '2011.09.15 全事業部を指定可能とする
                Select Case Last_JGYOBU
                    Case "*"
                
                        If StrConv(ITEMREC.JGYOBU, vbUnicode) = "S" Then
                            Exit Do
                        End If
                
                    Case Else
                
                        If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                            StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                            '事業部／国内外ブレーク
                            Exit Do
                        End If
                    
                        If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                        Else
                            If StrConv(ITEMREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                                '倉庫番号ブレーク
                                Exit Do
                            End If
                        End If
            
                End Select
    '2011.09.15 全事業部を指定可能とする
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        '-----------------------------------------  '商品化集計ファイル作成
        
In_Cnt = In_Cnt + 1
        
        
        
        
        
        '抽出条件変更       2011.07.04
        SKIP_FLG = False
        
    '2011.09.15 全事業部を指定可能とする
        If Last_JGYOBU = "*" Then
            
            If StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                SKIP_FLG = True
            End If
            
            
            If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
            Else
                If StrConv(ITEMREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                    '倉庫番号ブレーク
                    SKIP_FLG = True
                End If
            End If
        End If
    '2011.09.15 全事業部を指定可能とする
        
        
        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> GOODS_ON Then
            SKIP_FLG = True
        End If
                
        If StrConv(ITEMREC.GOODS_OUT_F, vbUnicode) = "1" Then
            SKIP_FLG = True
        End If
                
'2011.07.25
'        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "1" And StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "2" Then
        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "1" And StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "2" And Trim(StrConv(ITEMREC.NAI_BUHIN, vbUnicode)) <> "" Then
'2011.07.25
            
            
            SKIP_FLG = True
        End If
        
        If NOT_Hin_Name_F Then
            For i = 0 To UBound(NOT_Hin_Name)
                If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                    SKIP_FLG = True
                    Exit For
                End If
            Next i
        End If
        '抽出条件変更       2011.07.04
                
        
        
'2011.07.04        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
        If Not SKIP_FLG Then
                    
            
Out_Cnt = Out_Cnt + 1
            
            
            '2011.09.15
                                                    '事業部
''            Call UniCode_Conv(GOODSREC.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(GOODSREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            '2011.09.15
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
            
                                                    
            '2011.09.15
                                                    '在庫集計処理
'            If Zaiko_Syukei_Proc(Sumi_QTY, _
'                                    Mi_QTY, _
'                                    Last_JGYOBU, _
'                                    Right(Combo(pcmbNaigai).Text, 1), _
'                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
'                Exit Function
'            End If
            
'>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
            If Zaiko_Syukei_Proc(Sumi_QTY, _
                                    Mi_QTY, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    Right(Combo(pcmbNaigai).Text, 1), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                    , _
                                   GOODS_OFF_SOKO_NO, _
                                   True) = SYS_ERR Then
                Exit Function
            End If
            
            
'            If Zaiko_Syukei_Proc(Sumi_QTY, _
'                                    Mi_QTY, _
'                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                    Right(Combo(pcmbNaigai).Text, 1), _
'                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) _
'                                  = SYS_ERR Then
'                Exit Function
'            End If
'>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
            
            
            
            '2011.09.15
                                                    
                                                    
                                                    '商品化済み在庫数
            '2011.07.04
            Sumi_QTY = Sumi_QTY - SAMPLE_QTY
            If Sumi_QTY < 0 Then
                Sumi_QTY = 0
            End If
            '2011.07.04
            
            Call UniCode_Conv(GOODSREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                                    '未商品在庫数
            Call UniCode_Conv(GOODSREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                                    
                                                    
                                                    
            '2011.09.15
                                                    '月平均出荷数
''            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            '2011.09.15
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            AVE_QTY = 0
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
''2011.07.04                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
''2011.07.04                    AVE_QTY = CDbl(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY1, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY2, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, vbUnicode))
                    AVE_QTY = CDbl(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY1, "00000000")
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY2, "00000000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
                    Exit Function
            End Select
                                                    '事前商品化状況
            If AVE_QTY = 0 Then
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")
            Else
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
            End If
            
            
            
            '全事業部対応   2011.09.15
            If GetIni(App.EXEName, StrConv(ITEMREC.JGYOBU, vbUnicode), App.EXEName, c) Then
                SHIMUKE_CODE = ""
            Else
                SHIMUKE_CODE = Trim(c)
            End If
            '全事業部対応   2011.09.15
            
            
            
            
            '個装箱設定
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "010")
            sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(GOODSREC.KOSOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.KOSOU, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                    Exit Function
            End Select
            
            
            '外装箱設定
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "010")
            sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(GOODSREC.GAISOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Call UniCode_Conv(GOODSREC.KO_QTY, StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.GAISOU, "")
                    Call UniCode_Conv(GOODSREC.KO_QTY, "000.00")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                    Exit Function
            End Select
            
            
            '2011.07.04
            
            Call UniCode_Conv(GOODSREC.NAI_BUHIN, StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
            Call UniCode_Conv(GOODSREC.GAI_BUHIN, StrConv(ITEMREC.GAI_BUHIN, vbUnicode))
            '2011.07.04
            
            
            
            
            
            '2011.09.15 商品化用入荷予定獲得
            Call UniCode_Conv(GOODSREC.N_YOTEI_DT, "")
            Call UniCode_Conv(GOODSREC.N_YOTEI_QTY, "00000000")
            Call UniCode_Conv(GOODSREC.N_YOTEI_KEY_NO, "")
            
            
If Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) = "ARE03-630-HH" Then
    Debug.Print
End If
            
            
            Call UniCode_Conv(K1_PLN_Y_NYUKA.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.N_YOTEI_DT, Format(Text(ptxN_YOTEI_DT_S).Text, "YYYYMMDD"))
            Call UniCode_Conv(K1_PLN_Y_NYUKA.SEQ_NO, "")
            
            com = BtOpGetGreaterEqual
            
            Fast_Flg = False    '2011.12.16
            
            Do
                DoEvents
            
            
                sts = BTRV(com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K1_PLN_Y_NYUKA, Len(K1_PLN_Y_NYUKA), 1)
                Select Case sts
                    
                    Case BtNoErr
                        
                        
                        If StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
                            StrConv(PLN_Y_NYUKA_R.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
                            StrConv(PLN_Y_NYUKA_R.HIN_GAI, vbUnicode) <> StrConv(GOODSREC.HIN_GAI, vbUnicode) Then
                            
                            
                            If Not Fast_Flg Then
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
                            
                            Exit Do
                        End If
                    
                    
                        If StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode) > Format(Text(ptxN_YOTEI_DT_E).Text, "YYYYMMDD") Then
                            
                            If Not Fast_Flg Then
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
                            
                            Exit Do
                        End If
                    
                    
                    
                    Case BtErrEOF
                        
                        If Not Fast_Flg Then
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
                        
                        Exit Do
                    
                    Case Else
                        Call File_Error(sts, com, "商品化用入荷予定ファイル")
                        Exit Function
                End Select
            
                            
                Call UniCode_Conv(GOODSREC.N_YOTEI_DT, StrConv(PLN_Y_NYUKA_R.N_YOTEI_DT, vbUnicode))
                Call UniCode_Conv(GOODSREC.N_YOTEI_QTY, StrConv(PLN_Y_NYUKA_R.N_YOTEI_QTY, vbUnicode))
                
                Call UniCode_Conv(GOODSREC.N_YOTEI_KEY_NO, StrConv(PLN_Y_NYUKA_R.KEY_NO, vbUnicode))
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
'                Exit Do

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
            
                com = BtOpGetNext
                Fast_Flg = True
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
            
            
            Loop
            '2011.09.15
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            '2011.09.15
            
'Call LOG_OUT(LOG_F, StrConv(ITEMREC.JGYOBU, vbUnicode) & " " & StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode) & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
'            Do
'
'                sts = BTRV(BtOpInsert, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                        Beep
'                        ans = MsgBox("他端末でデータ使用中です。<GOODS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                        If ans = vbCancel Then
'                            Exit Function
'                        End If
'                    Case Else
'                        Call File_Error(sts, BtOpInsert, "商品化支援集計データ")
'                        Exit Function
'                End Select
'
'            Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.16
        End If
        
'全数表示とする　2011.09.26
If Right(Format(In_Cnt, "000"), 2) = "00" Then
    DoEvents
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "集計データ出力中！！[" & StrConv(GOODSREC.JGYOBU, vbUnicode) & "][" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)
    DoEvents
End If
        
        
        
        
        
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "集計データ出力中！！[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)
    DoEvents

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
Dim FileName        As String
    
''2011.07.04Dim wkSUMI_PERCENT      As Long



Dim SKIP_F              As Boolean
Dim FSW             As Boolean
    
    
'2011.01.13
Dim In_Cnt          As Long
Dim Out_Cnt         As Long

Dim JISEKI_TOTAL    As Double

Dim i               As Integer


Dim SHIJI_No        As String * 8
'2011.01.13
    
    
'2011.03.31
Dim MAIN_KOUTEI(0 To 9) _
                    As Long
Dim wkTANI          As Double
Dim wkQTY           As Double

Dim KOUSEI()        As KOUSEI_TBL
Dim j               As Integer
Dim KOUSEI_FLG      As Boolean

Dim wkInt           As Integer
'2011.03.31
    
    Data_Proc = True

In_Cnt = 0
Out_Cnt = 0


'2011.07.04
Dim Line_Cnt            As Long
Dim wkFROM_SUMI_PERCENT As Long
Dim wkTO_SUMI_PERCENT   As Long

Dim wkUKEHARAI_CODE     As String * 5

Dim wkHIN_NAME          As String * 40


Dim Out_Put_Rec         As String
'2011.07.04



'2011.09.15
Dim c                   As String * 128
'2011.09.15


Dim yn                  As Integer  '2012.05.08


Dim wkDate          As String * 8   '2012.08.27
Dim PN_JCode        As String       '2012.08.27
Dim ShisanJCode     As String       '2012.08.27

                            Dim chg As Integer


    Call Input_Lock
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    Command(7).Enabled = True
    Command(7).Caption = "中 断"
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08



    FileName = GOODS_DATA
    
    '2011.09.15 全事業部指定可
    If Last_JGYOBU = "*" Then
    Else
        sts = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), sts) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - sts)
    End If
    '2011.09.15 全事業部指定可
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (FileName) For Output As FileNo
    On Error GoTo 0


    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    Else
        If CANCEL_F = 9 Then
            Close #FileNo
            Data_Proc = False
            Exit Function
        End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    End If
    
    
''2011.07.04    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
''2011.07.04        wkSUMI_PERCENT = 100
''2011.07.04    Else
''2011.07.04        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
''2011.07.04    End If
    
    
    '2011.07.04
    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        wkFROM_SUMI_PERCENT = 0
    Else
        wkFROM_SUMI_PERCENT = CDbl(Text(ptxFROM_SUMI_PERCENT).Text)
    End If
    
    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        wkTO_SUMI_PERCENT = 999
    Else
        wkTO_SUMI_PERCENT = CDbl(Text(ptxTO_SUMI_PERCENT).Text)
    End If
    '2011.07.04
    
    
    
    
    
    
    FSW = True
    
    

    If SORT_SEQ = 0 Then        '2008.11.06


        '2011.09.15 全事業部指定可
        If Last_JGYOBU = "*" Then
        
            Call UniCode_Conv(K0_GOODS.JGYOBU, "")
            Call UniCode_Conv(K0_GOODS.NAIGAI, "")
            Call UniCode_Conv(K0_GOODS.ST_SOKO, "")
            Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
            Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
        Else
            Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K0_GOODS.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
            Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
        End If
        '2011.09.15 全事業部指定可
    
    Else
    
        '2011.09.15 全事業部指定可
        If Last_JGYOBU = "*" Then
            Call UniCode_Conv(K3_GOODS.JGYOBU, "")
            Call UniCode_Conv(K3_GOODS.NAIGAI, "")
            Call UniCode_Conv(K3_GOODS.ST_SOKO, "")
            Call UniCode_Conv(K3_GOODS.AVE_SYUKA, "zzzzzzzz")
            Call UniCode_Conv(K3_GOODS.Sumi_QTY, "")
            Call UniCode_Conv(K3_GOODS.Mi_QTY, "zzzzzzzz")
            Call UniCode_Conv(K3_GOODS.SUMI_PERCENT, "")
            Call UniCode_Conv(K3_GOODS.HIN_GAI, "")
        Else
            Call UniCode_Conv(K3_GOODS.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K3_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K3_GOODS.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K3_GOODS.AVE_SYUKA, "zzzzzzzz")
            Call UniCode_Conv(K3_GOODS.Sumi_QTY, "")
            Call UniCode_Conv(K3_GOODS.Mi_QTY, "zzzzzzzz")
            Call UniCode_Conv(K3_GOODS.SUMI_PERCENT, "")
            Call UniCode_Conv(K3_GOODS.HIN_GAI, "")
        End If
        '2011.09.15 全事業部指定可
    
    End If
    
    com = BtOpGetGreater
    
    
        
    '2011.07.04
    Line_Cnt = 0
    '2011.07.12
    In_Cnt = 0
    Out_Cnt = 0
    '2011.07.12
    
    Do
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.08
        DoEvents
        If CANCEL_F = 1 Then
            yn = MsgBox("中断します。よろしいですか？" & Chr(10) & "  はい：処理を中断します。" & Chr(10) & "いいえ：処理を継続します。", vbYesNo, "確認入力")
            If yn = vbYes Then
                Close #FileNo
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "CSVデータ出力中断！！", Me.hwnd, 0)
                Text1(1).Text = Now
                Call Input_UnLock
                Command(7).Enabled = True
                Command(7).Caption = "ﾃﾞｰﾀ"
                CANCEL_F = 9
                Data_Proc = False
                Exit Function
            Else
                CANCEL_F = 0
            End If
        End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.05.08
        
        If SORT_SEQ = 0 Then        '2008.11.06
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Else
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K3_GOODS, Len(K3_GOODS), 3)
        End If
        Select Case sts
            Case BtNoErr
                
If Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) = "AMX03B-R60" Then
Debug.Print
End If
                
                
                '2011.09.15 全事業部指定可
                If Last_JGYOBU = "*" Then
                Else
                    If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                                            
                                            
                                            
                                            
                    
                    If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                    Else
                        If StrConv(GOODSREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                            Exit Do
                        End If
                    End If
                End If
                '2011.09.15 全事業部指定可
                                
                
                
                SKIP_F = False
                
                
                
                '2011.09.15 全事業部指定可
                If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Else
                    If StrConv(GOODSREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                        SKIP_F = True
                    End If
                End If
                '2011.09.15 全事業部指定可
                
                
                
                
                
                
                
                If Not IsNumeric(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
''2011.07.04                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
''2011.07.04                        SKIP_F = True
''2011.07.04                    End If
                
''2011.07.04
                    If Val(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) < wkFROM_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
                    
                    If Val(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkTO_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
''2011.07.04
                End If


                
                '2011.09.15
'                If Val(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
'                    SKIP_F = True
'                End If
                '2011.09.15
                
                
                                
                '>>>>>>>>>>>>>>>>>>>>>>>    2012.05.09
                'If Val(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 And Val(StrConv(GOODSREC.N_YOTEI_QTY, vbUnicode)) <= 0 Then
                '    SKIP_F = True
                'End If
                
                
                
                '未商品=０　かつ　月平均出荷数＝０　は　出力しない
                If Val(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 And Val(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) <= 0 Then
                    SKIP_F = True
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>    2012.05.09


                





            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select

        
'-------------------------------------------------  明細印刷
        
        In_Cnt = In_Cnt + 1
        
        
        If Not SKIP_F Then
        
        


        
        
        
            If FSW Then
                
                FSW = False
                        
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
''2011.07.12                Line_Cnt = Line_Cnt + 1
''2011.07.12                Write #FileNo, "*** 商品化支援アラームリスト　***"
''2011.07.04                Write #FileNo, "作成日付:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                        
            
'                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況"
''2011.01.13                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況", "工数　（分／個）"
                
''2011.03.31                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況", "見積工数　（分／個）", "実績工数　（分／個）"
                
                
''2011.07.04                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況", "見積工数　（分／個）", "実績工数　（分／個）", "作業工数"
                
                
                '2011.07.04
                Line_Cnt = Line_Cnt + 1
''2011.07.12                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化予定数", "外装品番", "外装使用枚数", "商品化済在庫", "未商品在庫", "未商品　別置1", "未商品　別置2", TUKI1_TITLE, S_TUKI1_TITLE, S_TUKI2_TITLE, "事前商品化必要数", "事前商品化状況", "見積工数　（分／個）", "実績工数　（分／個）", "作業工数", "国内供給区分", "海外供給区分", "商品化完了手配先"
                '2011.07.12
                
                
                
''2011.09.15                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化予定数", "商品化予定工数", "外装品番", "外装使用枚数", "商品化済在庫", "未商品在庫", "未商品　別置1", "未商品　別置在庫1", "未商品　別置2", "未商品　別置在庫2", TUKI1_TITLE, S_TUKI1_TITLE, S_TUKI2_TITLE, "事前商品化必要数", "事前商品化状況", "見積工数　（分／個）", "実績工数　（分／個）", "作業工数", "国内供給区分", "海外供給区分", "商品化完了手配先"
                
                
            
''2011.07.04                Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
                '2011.09.15
'2012.05.08                Write #FileNo, "BU区分", "標準棚番", "品番（外部）", "商品化予定数", "商品化済在庫", "未商品在庫", S_TUKI1_TITLE, "事前商品化状況", "部品入荷予定日", "部品入荷予定数", "見積工数　（分／個）", "商品化予定工数", "商品化予定日", "資材（箱№）", "外装品番", "外装使用枚数", "未商品　別置1", "未商品　別置在庫1", "未商品　別置2", "未商品　別置在庫2", "事前商品化必要数", "実績工数　（分／個）", "作業工数", "国内供給区分", "海外供給区分", "商品化完了手配先"
                '2012.05.08
'2012.05.10                Write #FileNo, "BU区分", "標準棚番", "品番（外部）", "商品化予定数", "商品化済在庫", "未商品在庫", S_TUKI1_TITLE, "事前商品化状況", "部品入荷予定日", "部品入荷予定数", "見積工数　（分／個）", "商品化予定工数", "商品化予定日", "資材（箱№）", "外装品番", "外装使用枚数", "未商品　別置1", "未商品　別置在庫1", "未商品　別置2", "未商品　別置在庫2", "事前商品化必要数", "実績工数　（分／個）", "商品化工数", "国内供給区分", "海外供給区分", "商品化完了手配先"
                
                '2012.05.10
'2012.08.27                Write #FileNo, "BU区分", "標準棚番", "品番（外部）", "商品化予定数", "商品化済在庫", "未商品在庫", S_TUKI1_TITLE, "事前商品化状況", "部品入荷予定日", "部品入荷予定数", "商品化予定工数", "商品化予定日", "資材（箱№）", "未商品　別置1", "未商品　別置在庫1", "未商品　別置2", "未商品　別置在庫2", "商品化工数", "事前商品化必要数", "実績工数　（分／個）", "見積工数　（分／個）", "外装品番", "外装使用枚数", "国内供給区分", "海外供給区分", "商品化完了手配先"
            
                '2012.08.27
                Write #FileNo, "BU区分", "標準棚番", "品番（外部）", "商品化予定数", "商品化済在庫", "未商品在庫", S_TUKI1_TITLE, "事前商品化状況", "部品入荷予定日", "部品入荷予定数", "商品化予定工数", "商品化予定日", "資材（箱№）", "未商品　別置1", "未商品　別置在庫1", "未商品　別置2", "未商品　別置在庫2", "商品化工数", "事前商品化必要数", "実績工数　（分／個）", "見積工数　（分／個）", "外装品番", "外装使用枚数", "国内供給区分", "海外供給区分", "国内供給打切年月", "商品化完了手配先"
            
            
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

''2011.07.04                Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"


            End If
            
            If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            
                '2011.09.15
                            '設定発注点より大きい
                'Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
                'Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzz")
                'com = BtOpGetGreaterEqual
                '2011.09.15
            Else
                
                If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Then
                Else
                    Line_Cnt = Line_Cnt + 1
                
                
                                                            
''''''''''''''''''''    2011.09.15  項目追加／削除　位置変更
                                                            
                                                            'BU区分
                    Write #FileNo, StrConv(GOODSREC.JGYOBU, vbUnicode),
                                                            '標準棚番
                    Edit = StrConv(GOODSREC.ST_SOKO, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
                    Write #FileNo, Edit,
                                                            '品番（外部）
    
                    Write #FileNo, Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)),
                                                            
                                                            '商品化予定数(入力項目)     2011.07.12
                    Write #FileNo, "",                                                  '2011.07.04
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
                                                            '生産計画月平均出荷数(1)    2011.07.04
                    Edit = Format(CDbl(StrConv(GOODSREC.S_AVE_SYUKA_QTY1, vbUnicode)), "#,##0.0")
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, "=""" & Edit & """",
                                                            '事前商品化状況
                    Edit = Format(CInt(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            '部品入荷予定日 2011.09.15
                    If Trim(StrConv(GOODSREC.N_YOTEI_DT, vbUnicode)) = "" Then
                        Write #FileNo, "",
                    Else
                        Write #FileNo, Mid(StrConv(GOODSREC.N_YOTEI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(GOODSREC.N_YOTEI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(GOODSREC.N_YOTEI_DT, vbUnicode), 7, 2),
                    End If
                                                            '部品入荷予定数 2011.09.15
                    Edit = Format(Val(StrConv(GOODSREC.N_YOTEI_QTY, vbUnicode)), "#")
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            
                                                            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                                                            
                                                            '見積工数
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.03.31
'                    For i = 0 To UBound(MAIN_KOUTEI)
'                        MAIN_KOUTEI(i) = 0
'                    Next i
'
'
'                    '①
'                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
'
'                        wkTANI = Val(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode))
'                    Else
'                        wkTANI = 0
'                    End If
'                    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
'                        '2009.09.18
'                        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
'                            wkQTY = Val(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode))
'                        Else
'                            wkQTY = 1
'                        End If
'                    Else
'                        wkQTY = 1
'                    End If
'                    MAIN_KOUTEI(0) = wkTANI * wkQTY
'
'
'
'
'
'                    '②
'                    '-------------------　構成情報テーブル展開
'
'                    Erase KOUSEI
'                    i = -1
'
'                    KOUSEI_FLG = False
'
'
'                    '全事業部対応   2011.09.15
'                    If GetIni(App.EXEName, StrConv(GOODSREC.JGYOBU, vbUnicode), App.EXEName, c) Then
'                        SHIMUKE_CODE = ""
'                    Else
'                        SHIMUKE_CODE = Trim(c)
'                    End If
'                    '全事業部対応   2011.09.15
'
'
'
'                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
'                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
'                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
'                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
'
'                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
'                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
'
'                    com = BtOpGetGreater
'
'                    Do
'                        DoEvents
'
'                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
'                        Select Case sts
'                            Case BtNoErr
'
'
'                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
'                                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
'                                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
'                                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
'
'                                    Exit Do
'
'                                End If
'
'                            Case BtErrEOF
'                                Exit Do
'                            Case Else
'                                Call File_Error(sts, BtOpGetNext, "構成マスタ")
'                                Exit Function
'                        End Select
'
'                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
'                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
'                        End If
'                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
'                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
'                        End If
'
'                        i = i + 1
'                        KOUSEI_FLG = True
'
'                        ReDim Preserve KOUSEI(0 To i)
'                        '事業部
'                        KOUSEI(i).KO_JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
'                        '国内外
'                        KOUSEI(i).KO_NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
'
'                        '種別
'                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
'                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
'                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'
'                        Select Case sts
'                            Case BtNoErr
'                                KOUSEI(i).KO_SYUBETSU = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
'                            Case BtErrKeyNotFound
'                                KOUSEI(i).KO_SYUBETSU = ""
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
'                                Exit Function
'
'                        End Select
'
'                        '品番
'                        KOUSEI(i).KO_HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'
'                            Case BtErrKeyNotFound
'
'                                Call UniCode_Conv(ITEMREC.SEI_KBN, "")
'                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
'                                Call UniCode_Conv(ITEMREC.S_KOUSU, "")
'                                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
'
'
'
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                                Exit Function
'                        End Select
'
'                        '員数
'                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
'                            KOUSEI(i).KO_QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
'                        Else
'                            KOUSEI(i).KO_QTY = 1#
'                        End If
'                        '仕入単価
'                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
'                            KOUSEI(i).G_ST_SHITAN = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
'                        Else
'                            KOUSEI(i).G_ST_SHITAN = 0#
'                       End If
'
'                        '売上単価
'                        Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
'
'                            Case "1"
'                               KOUSEI(i).G_ST_URITAN = 0#
'                            Case "2"
'                               KOUSEI(i).G_ST_URITAN = 0#
'                            Case Else
'                                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
'                                    KOUSEI(i).G_ST_URITAN = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
'                                Else
'                                    KOUSEI(i).G_ST_URITAN = 0#
'                                End If
'                        End Select
'                        '仕入金額計
'                        KOUSEI(i).G_ST_SHIKIN = 0#
'                        For j = 0 To UBound(SHIZAI_T)
'                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
'
'
'                                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
'
'                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
'
'                                        If CDbl(KOUSEI(i).KO_QTY) = 0 Then '2010.02.22
'                                            KOUSEI(i).G_ST_SHIKIN = 0#
'                                        Else
'                                            KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_SHITAN)) / CDbl(KOUSEI(i).KO_QTY), 2)
'                                        End If
'                                    Else
'                                        KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_SHITAN), 2)
'                                    End If
'                                End If
'                                Exit For
'                            End If
'
'                        Next j
'                       '売上金額計
'                        KOUSEI(i).G_ST_URIKIN = 0
'                        KOUSEI(i).G_ST_URIKIN_KUSATU = 0
'
'                        For j = 0 To UBound(SHIZAI_T)
'
'                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
'
'
'                                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
'                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
'
'
'
'                                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
'                                            KOUSEI(i).G_ST_URIKIN = 0#
'                                        Else
'                                            KOUSEI(i).G_ST_URIKIN = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
'                                        End If
'                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN), 2)
'                                    Else
'                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY) * CDbl(KOUSEI(i).G_ST_URITAN)), 2)
'                                    End If
'
'
'                                Else
'
'                                    If KUSATU_F Then
'
'                                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
'
'
'                                            If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
'                                                KOUSEI(i).G_ST_URIKIN_KUSATU = 0
'                                            Else
'                                                KOUSEI(i).G_ST_URIKIN_KUSATU = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
'                                            End If
'                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN_KUSATU), 2)
'
'                                        Else
'                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_URITAN), 2)
'                                        End If
'
'
'                                    End If
'                                End If
'                            End If
'                        Next j
'
'                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
'                            KOUSEI(i).S_KOUSU = 0
'                            KOUSEI(i).SEI_SYU_KON = 0
'                        Else
'                            '作業時間
'                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
'                                KOUSEI(i).S_KOUSU = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
'                            Else
'                                KOUSEI(i).S_KOUSU = 0#
'                            End If
'                            '集合梱包
'                            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
'                                KOUSEI(i).SEI_SYU_KON = CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode))
'                            Else
'                                KOUSEI(i).SEI_SYU_KON = 0#
'                            End If
'                        End If
'                    Loop
'
'
'                    wkInt = 0
'                    If KOUSEI_FLG Then
'                        For i = 0 To UBound(KOUSEI)
'
'                            For j = 0 To UBound(SHIZAI_T)
'                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
'                                    wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).S_KOUSU) * CDbl(KOUSEI(i).KO_QTY), 0))
'                                    Exit For
'                                End If
'
'                            Next j
'
'                        Next i
'                    End If
'
'                    wkTANI = wkInt
'                    wkQTY = 1
'                    MAIN_KOUTEI(1) = wkTANI * wkQTY
'
'                    '③
'                    wkInt = 0
'                    If KOUSEI_FLG Then
'                        For i = 0 To UBound(KOUSEI)
'
'                            For j = 0 To UBound(DOUKON_T)
'                                If KOUSEI(i).KO_SYUBETSU = DOUKON_T(j) Then
'
'                                    If IsNumeric(KOUSEI(i).KO_QTY) Then
'                                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).KO_QTY), 0))
'                                    End If
'
'
'
'                                    Exit For
'                                End If
'
'                            Next j
'
'                        Next i
'                    End If
'                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
'                        wkTANI = CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode))
'                    Else
'                        wkTANI = 0#
'                    End If
'                    wkQTY = wkInt
'                    MAIN_KOUTEI(2) = wkTANI * wkQTY
'
'
'
'
'                    '④
'                    wkInt = 0
'                    If KOUSEI_FLG Then
'                        For i = 0 To UBound(KOUSEI)
'
'                            For j = 0 To UBound(KAKOU_T)
'                                If KOUSEI(i).KO_SYUBETSU = KAKOU_T(j) Then
'                                    If IsNumeric(KOUSEI(i).S_KOUSU) Then
'                                        wkInt = wkInt + CInt(KOUSEI(i).S_KOUSU)
'                                    End If
'                                    Exit For
'                                End If
'
'                            Next j
'
'                        Next i
'                    End If
'                    wkTANI = wkInt
'                    wkQTY = 1
'                    MAIN_KOUTEI(3) = wkTANI * wkQTY
'
'                    '⑤
'                    wkInt = 0
'                    If KOUSEI_FLG Then
'                        For i = 0 To UBound(KOUSEI)
'
'
'                            For j = 0 To UBound(SHIZAI_T)
'
'                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
'                                    If IsNumeric(KOUSEI(i).SEI_SYU_KON) Then
'                                        wkInt = wkInt + CInt(KOUSEI(i).SEI_SYU_KON)
'                                    End If
'                                End If
'
'                            Next j
'
'                        Next i
'                    End If
'                    wkTANI = wkInt
'                    wkQTY = 1
'                    MAIN_KOUTEI(4) = wkTANI * wkQTY
'
'
'                    '計
'                    wkInt = 0
'                    For i = 0 To UBound(MAIN_KOUTEI)
'
'                        wkInt = wkInt + MAIN_KOUTEI(i)
'                    Next i
'                    Edit = Format(ToHalfAdjust(CCur(wkInt) / 60, 1), "#0.0")
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.12.22
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.12.22
'                    If Len(Edit) < 8 Then
'                        Edit = Space(8 - Len(Edit)) & Edit
'                    End If
'
'                    Write #FileNo, "=""" & Edit & """",
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.03.31
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    
                    
                    
                    
                    
                    
                    
                    '商品化予定工数             2011.07.12
'2012.05.09                    Write #FileNo, "=round(D" & Format(Line_Cnt, "#") & "*K" & Format(Line_Cnt, "#") & ",1)",
'2012.05.10                    Write #FileNo, "=round(D" & Format(Line_Cnt, "#") & "*W" & Format(Line_Cnt, "#") & ",1)",
                    
                    Write #FileNo, "=round(D" & Format(Line_Cnt, "#") & "*R" & Format(Line_Cnt, "#") & ",1)",           '2012.05.10
                    '商品化予定日               2011.09.15
                    Write #FileNo, "",
                                                            '箱№
                    'Write #FileNo, Trim(StrConv(GOODSREC.PACKING_NO, vbUnicode)),      '2008.03.03
                    Write #FileNo, Trim(StrConv(GOODSREC.KOSOU, vbUnicode)),            '2008.03.03
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
'                    '外装箱 2011.07.04
'                    Write #FileNo, Trim(StrConv(GOODSREC.GAISOU, vbUnicode)),
'                    '外装箱使用枚数 2011.07.04
'                    If Val(StrConv(GOODSREC.KO_QTY, vbUnicode)) = 0 Then
'                        Write #FileNo, 0,
'                    Else
'                        Write #FileNo, "=roundup(D" & Format(Line_Cnt, "#") & "/" & CDbl(StrConv(GOODSREC.KO_QTY, vbUnicode)) & ",0)",
'                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    
                    If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.JGYOBU, vbUnicode), StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
                                                            '未商品別置き
                    If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) = 0 Then
''2011.07.12                        Write #FileNo, ,
                        '2011.07.12
                        Write #FileNo, , ,
                    Else
                        Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(0).EE_LOC, 2)
                        
                        '2011.07.12
                        Write #FileNo, Edit,
                        Edit = ""
                        Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                        '2011.07.12
                        Write #FileNo, Edit,
                    End If
                                                            
                    If Len(Trim(EE_ZAIKO_TBL(1).EE_LOC)) = 0 Then
''2011.07.12                        Write #FileNo, ,
                        Write #FileNo, , ,
                    Else
                        Edit = Left(EE_ZAIKO_TBL(1).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(1).EE_LOC, 2)
                        
                        
                        '2011.07.12
                        Write #FileNo, Edit,
                        Edit = ""
                        Edit = Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                        '2011.07.12
                        
                        
''2011.07.12                        Edit = Edit & " " & Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                        Write #FileNo, Edit,
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    '作業工数　2008.09.19
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                    
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                            
                            
                            
                                Edit = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.0")
                            
                            
                            
                            


                            
                            
                            
                                If Len(Edit) < 8 Then
                                    Edit = Space(8 - Len(Edit)) & Edit
                                End If
                            
                            
                            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.09.24
                                If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) And IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
                                    If CDbl(StrConv(ITEMREC.SEI_RATE, vbUnicode)) <> 0 Then
                                
                                        Edit = Format(Round(CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) / CDbl(StrConv(ITEMREC.SEI_RATE, vbUnicode)), 2), "#0.0")
                                    
'2018.08.06Call LOG_OUT(LOG_F, StrConv(ITEMREC.JGYOBU, vbUnicode) & " " & StrConv(ITEMREC.NAIGAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                    Else
                                        Edit = "0.0"
                                    End If
                                Else
                                
                                    Edit = "0.0"
                                End If
                            
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.09.24
                                If Len(Edit) < 8 Then
                                    Edit = Space(8 - Len(Edit)) & Edit
                                End If
                            
                            
                            End If
                            wkHIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            Write #FileNo, "=""" & Edit & """",
                        
                        
                                                    
                        Case BtErrKeyNotFound
                
                            Edit = ""
                            
                            wkHIN_NAME = ""
                            
                            
                                                
                        
                            Write #FileNo, Edit,
                        
                            '2011.07.04
                            Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")
                            '2011.07.04
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    '2011.07.04
                                                            '事前商品化必要数
                    'Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    Edit = Format(CLng(StrConv(GOODSREC.S_AVE_SYUKA_QTY1, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    '2011.07.04
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            
                    '2011.01.13 実績工数
                    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_F, P_KAN_ON)
                    Call UniCode_Conv(K1_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
                    Call UniCode_Conv(K1_P_SSHIJI_O.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_DT, "zzzzzzzz")
                    Call UniCode_Conv(K1_P_SSHIJI_O.SHIJI_No, "zzzzzzzz")
                    
                    JISEKI_TOTAL = 0
                    
                    SHIJI_No = ""
                                    
                    Do
                    
                    
                    
                        DoEvents
                        wkUKEHARAI_CODE = ""
                        sts = BTRV(BtOpGetLess, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K1_P_SSHIJI_O, Len(K1_P_SSHIJI_O), 1)
                        Select Case sts
                            Case BtNoErr
                                If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Or _
                                    StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
                                    StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
                                    StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
                                    StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> StrConv(GOODSREC.HIN_GAI, vbUnicode) Then
                                    Exit Do
                                Else
                                    
                                    
                                    
                                    If Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <= 0 Then
                                    Else
                                    
                                        
                                        SHIJI_No = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)
                                                        
                                        JISEKI_TOTAL = 0
                            
                            Debug.Print StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)
                                        
                                        For i = 0 To 9
                                        
                                            If IsNumeric(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) And IsNumeric(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)) Then
                                                JISEKI_TOTAL = JISEKI_TOTAL + Round(Val(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * Val(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)), 2)
                                            End If
                                        
                                        Next i
                                    
                                        wkUKEHARAI_CODE = StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)
                                    
                                        Exit Do
                                    End If
                                                                    
                                End If
                        
                            Case BtErrEOF
                                Exit Do
                            Case Else
                        
                                Call File_Error(sts, BtOpGetLess, "指図票データ(親)")
                                Exit Function
                        End Select
                    
                    Loop
                    
                    Edit = ""
                    If IsNumeric(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) Then
                        If Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                            Edit = Format(Round(CDbl(JISEKI_TOTAL / Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 2), "#0.0")
                        End If
                    End If
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, "=""" & Edit & """",
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                                                            '見積工数
                    For i = 0 To UBound(MAIN_KOUTEI)
                        MAIN_KOUTEI(i) = 0
                    Next i
                                        
                                        
                    '①
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
                        
                        wkTANI = Val(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0
                    End If
                    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                        '2009.09.18
                        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
                            wkQTY = Val(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode))
                        Else
                            wkQTY = 1
                        End If
                    Else
                        wkQTY = 1
                    End If
                    MAIN_KOUTEI(0) = wkTANI * wkQTY





                    '②
                    '-------------------　構成情報テーブル展開
                        
                    Erase KOUSEI
                    i = -1
        
                    KOUSEI_FLG = False
                                    
                                    
                    '全事業部対応   2011.09.15
                    If GetIni(App.EXEName, StrConv(GOODSREC.JGYOBU, vbUnicode), App.EXEName, c) Then
                        SHIMUKE_CODE = ""
                    Else
                        SHIMUKE_CODE = Trim(c)
                    End If
                    '全事業部対応   2011.09.15
                                    
                                    
                            
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                       
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
                                    
                    com = BtOpGetGreater
                                    
                    Do
                        DoEvents
                
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                
                                
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
                                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
                                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
                                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                                
                                    Exit Do
                            
                                End If
                
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                                Exit Function
                        End Select
                
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
                        End If
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
                        End If
                
                        i = i + 1
                        KOUSEI_FLG = True
                            
                        ReDim Preserve KOUSEI(0 To i)
                        '事業部
                        KOUSEI(i).KO_JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                        '国内外
                        KOUSEI(i).KO_NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                        
                        '種別
                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            
                        Select Case sts
                            Case BtNoErr
                                KOUSEI(i).KO_SYUBETSU = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
                            Case BtErrKeyNotFound
                                KOUSEI(i).KO_SYUBETSU = ""
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                                Exit Function
                        
                        End Select
                        
                        '品番
                        KOUSEI(i).KO_HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                         
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                                    
                                Call UniCode_Conv(ITEMREC.SEI_KBN, "")
                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
                                Call UniCode_Conv(ITEMREC.S_KOUSU, "")
                                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
                            
                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        '員数
                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                            KOUSEI(i).KO_QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                        Else
                            KOUSEI(i).KO_QTY = 1#
                        End If
                        '仕入単価
                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                            KOUSEI(i).G_ST_SHITAN = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
                        Else
                            KOUSEI(i).G_ST_SHITAN = 0#
                        End If
                    
                        '売上単価
                        Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                        
                            Case "1"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case "2"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case Else
                                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                    KOUSEI(i).G_ST_URITAN = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                                Else
                                    KOUSEI(i).G_ST_URITAN = 0#
                                End If
                        End Select
                        '仕入金額計
                        KOUSEI(i).G_ST_SHIKIN = 0#
                        For j = 0 To UBound(SHIZAI_T)
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                
                                
                                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                                    
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        If CDbl(KOUSEI(i).KO_QTY) = 0 Then '2010.02.22
                                            KOUSEI(i).G_ST_SHIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_SHITAN)) / CDbl(KOUSEI(i).KO_QTY), 2)
                                        End If
                                    Else
                                        KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_SHITAN), 2)
                                    End If
                                End If
                                Exit For
                            End If
                        
                        Next j
                       '売上金額計
                        KOUSEI(i).G_ST_URIKIN = 0
                        KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                
                        For j = 0 To UBound(SHIZAI_T)
                       
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                       
                       
                                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                    
                                    
                                    
                                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                            KOUSEI(i).G_ST_URIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_URIKIN = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                        End If
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN), 2)
                                    Else
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY) * CDbl(KOUSEI(i).G_ST_URITAN)), 2)
                                    End If
                        
                                
                                Else
                               
                                    If KUSATU_F Then
                                
                                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        
                                            If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                                            Else
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                            End If
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN_KUSATU), 2)
                                        
                                        Else
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_URITAN), 2)
                                        End If
                                    
                                    
                                    End If
                                End If
                            End If
                        Next j
                        
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                            KOUSEI(i).S_KOUSU = 0
                            KOUSEI(i).SEI_SYU_KON = 0
                        Else
                            '作業時間
                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                                KOUSEI(i).S_KOUSU = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
                            Else
                                KOUSEI(i).S_KOUSU = 0#
                            End If
                            '集合梱包
                            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                                KOUSEI(i).SEI_SYU_KON = CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode))
                            Else
                                KOUSEI(i).SEI_SYU_KON = 0#
                            End If
                        End If
                    Loop


                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(SHIZAI_T)
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).S_KOUSU) * CDbl(KOUSEI(i).KO_QTY), 0))
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(1) = wkTANI * wkQTY

                    '③
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(DOUKON_T)
                                If KOUSEI(i).KO_SYUBETSU = DOUKON_T(j) Then
                                    
                                    If IsNumeric(KOUSEI(i).KO_QTY) Then
                                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).KO_QTY), 0))
                                    End If
                                    
                                    
                                    
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
                        wkTANI = CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0#
                    End If
                    wkQTY = wkInt
                    MAIN_KOUTEI(2) = wkTANI * wkQTY




                    '④
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(KAKOU_T)
                                If KOUSEI(i).KO_SYUBETSU = KAKOU_T(j) Then
                                    If IsNumeric(KOUSEI(i).S_KOUSU) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).S_KOUSU)
                                    End If
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(3) = wkTANI * wkQTY
                    
                    '⑤
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                            
                            
                            For j = 0 To UBound(SHIZAI_T)
                            
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    If IsNumeric(KOUSEI(i).SEI_SYU_KON) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).SEI_SYU_KON)
                                    End If
                                End If
                            
                            Next j
                            
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(4) = wkTANI * wkQTY


                    '計
                    wkInt = 0
                    For i = 0 To UBound(MAIN_KOUTEI)
                    
                        wkInt = wkInt + MAIN_KOUTEI(i)
                    Next i
                    Edit = Format(ToHalfAdjust(CCur(wkInt) / 60, 1), "#0.0")
                    
                    
                    If Len(Edit) < 8 Then
                        Edit = Space(8 - Len(Edit)) & Edit
                    End If
                    
                    Write #FileNo, "=""" & Edit & """",

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    '外装箱 2011.07.04
                    Write #FileNo, Trim(StrConv(GOODSREC.GAISOU, vbUnicode)),
                    '外装箱使用枚数 2011.07.04
                    If Val(StrConv(GOODSREC.KO_QTY, vbUnicode)) = 0 Then
                        Write #FileNo, 0,
                    Else
                        Write #FileNo, "=roundup(D" & Format(Line_Cnt, "#") & "/" & CDbl(StrConv(GOODSREC.KO_QTY, vbUnicode)) & ",0)",
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    
                    
                    
                    
                    
                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
'                    '作業工数　2008.09.19
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
'
'
'                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
'
'
'
'                                Edit = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.0")
'
'
'                                If Len(Edit) < 8 Then
'                                    Edit = Space(8 - Len(Edit)) & Edit
'                                End If
'
'
'                            Else
'                                Edit = "0.0"
'
'                                If Len(Edit) < 8 Then
'                                    Edit = Space(8 - Len(Edit)) & Edit
'                                End If
'
'
'                            End If
'                            wkHIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'                            Write #FileNo, "=""" & Edit & """",
'
'
'
'                        Case BtErrKeyNotFound
'
'                            Edit = ""
'
'                            wkHIN_NAME = ""
'
'
'
'
'                            Write #FileNo, Edit,
'
'                            '2011.07.04
'                            Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")
'                            '2011.07.04
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                            Exit Function
'                    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.10
                    '国内供給区分   2011.07.04
                    Write #FileNo, StrConv(GOODSREC.NAI_BUHIN, vbUnicode),
                    '海外供給区分   2011.07.04
                    Write #FileNo, StrConv(ITEMREC.NAI_BUHIN, vbUnicode),
                    
                    
                    
                    
                    
                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.08.27
                    sts = GetIni("JCODE", StrConv(GOODSREC.JGYOBU, vbUnicode), "PN_JCode", c)
                    If sts Then
                        c = ""
                    End If
                    PN_JCode = Trim(c)
                    
                    sts = GetIni("ShisanJCode", StrConv(GOODSREC.JGYOBU, vbUnicode), "PN_JCode", c)
                    If sts Then
                        c = ""
                    End If
                    ShisanJCode = Trim(c)
                    
                    Call UniCode_Conv(K0_PN_M.JCode, PN_JCode)
                    Call UniCode_Conv(K0_PN_M.ShisanJCode, ShisanJCode)
                    Call UniCode_Conv(K0_PN_M.PN, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                                
                                
                    sts = BTRV(BtOpGetEqual, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
                    Select Case sts
                        Case BtNoErr
                            
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(PN_MREC.NaiDisconYm, "")
                        
                                                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "PNﾏｽﾀ")
                            Exit Function
                    End Select


                    wkDate = Trim(StrConv(PN_MREC.NaiDisconYm, vbUnicode)) & "01"
                    
                    If Not IsDate(Mid(wkDate, 1, 4) & "/" & Mid(wkDate, 5, 2) & "/" & Mid(wkDate, 7, 2)) Then
                        
                        
                        Write #FileNo, "=""" & Trim(StrConv(PN_MREC.NaiDisconYm, vbUnicode)) & """",
                    Else
                    
                        For i = 31 To 28 Step -1
                        
                            If IsDate(Mid(wkDate, 1, 4) & "/" & Mid(wkDate, 5, 2) & "/" & Format(i, "00")) Then
                                Exit For
                            End If
                        
                        Next i
                        Write #FileNo, Mid(wkDate, 1, 4) & "/" & Mid(wkDate, 5, 2) & "/" & Format(i, "00"),
                    
                    
                    End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.08.27
                    
                    
                    
                    
                    
                    
                    
                    '商品化完了手配先   2011.07.04
                    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, wkUKEHARAI_CODE)
                    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Select Case sts
                    Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                            Exit Function
                    End Select
                    Write #FileNo, wkUKEHARAI_CODE & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode)),
                    Write #FileNo,
                                                            
                                                            
''2011.07.04                    If Len(Trim(EE_ZAIKO_TBL(2).EE_LOC)) = 0 Then
''2011.07.04                        Write #FileNo, ,
''2011.07.04                    Else
''2011.07.04                        Edit = Left(EE_ZAIKO_TBL(2).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(2).EE_LOC, 2)
''2011.07.04                        Edit = Edit & " " & Format(EE_ZAIKO_TBL(2).EE_QTY, "#0")
''2011.07.04                        Write #FileNo, Edit,
''2011.07.04                    End If
                                                            
                                                            '月平均出荷数
'2011.09.15                    Edit = Format(CDbl(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0.0")
'2011.09.15                    If Len(Edit) < 8 Then
'2011.09.15                        Edit = Space(8 - Len(Edit)) & Edit
'2011.09.15                    End If
'2011.09.15                    Write #FileNo, "=""" & Edit & """",
'2011.09.15
                                                            '生産計画月平均出荷数(2)    2011.07.04
'2011.09.15                    Edit = Format(CDbl(StrConv(GOODSREC.S_AVE_SYUKA_QTY2, vbUnicode)), "#,##0.0")
'2011.09.15                    If Len(Edit) < 8 Then
'2011.09.15                        Edit = Space(8 - Len(Edit)) & Edit
'2011.09.15                    End If
'2011.09.15                    Write #FileNo, "=""" & Edit & """",
''''''''''''''''''''    2011.09.15  項目追加／削除　位置変更
                    
                    
                    
                    '2011.07.12
                    Out_Cnt = Out_Cnt + 1
                    
                    
                End If
            End If
            
'全数表示とする 2011.09.26
If Right(Format(In_Cnt, "000"), 2) = "00" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "CSVデータ出力中！！[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)
    DoEvents
End If
            
            com = BtOpGetNext
        End If
    Loop

    Close #FileNo

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "CSVデータ出力完了！！[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)

Text1(1).Text = Now


    Beep
    DoEvents
    MsgBox "「" & FileName & "」は正常に出力されました。"

    Call Input_UnLock
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    Command(7).Enabled = True
    Command(7).Caption = "ﾃﾞｰﾀ"
    CANCEL_F = 9
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
    
    
    
    Data_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
        Command(7).Enabled = True
        Command(7).Caption = "ﾃﾞｰﾀ"
        CANCEL_F = 9
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.05.08
        
        
        
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Proc = True
    End If


End Function

Private Function MI_ZAIKO_KENSAKU(JGYOBU As String, Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   未商品の処理
'----------------------------------------------------------------------------
Dim i           As Integer
Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long

Dim com         As Integer
Dim sts         As Integer

Dim GOODS_OFF_SOKO_FLG   As Integer  '2012.02.24

    MI_ZAIKO_KENSAKU = True
    
    For i = 0 To UBound(EE_ZAIKO_TBL)
        EE_ZAIKO_TBL(i).EE_LOC = ""
        EE_ZAIKO_TBL(i).EE_QTY = 0
    Next i
    
    '2011.09.15
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, JGYOBU)
    '2011.09.15
    
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
    
    '2012.02.24
'    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_ON)
    '2012.02.24
    
    
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
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
                
                '2012.02.24
                'If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
                '    Exit Do
                'End If
                GOODS_OFF_SOKO_FLG = False
                For i = 0 To UBound(GOODS_OFF_SOKO_NO)
                    If GOODS_OFF_SOKO_NO(i) = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                        GOODS_OFF_SOKO_FLG = True
                        Exit For
                    End If
                Next i
                '2012.02.24
                
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select
        
        
        '2012.02.24 IF分追加
        If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF And GOODS_OFF_SOKO_FLG = False Then
        Else
            For i = 0 To UBound(EE_ZAIKO_TBL)
                            
                If Trim(EE_ZAIKO_TBL(i).EE_LOC) = Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
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
        End If
        '2012.02.24 IF分追加
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function
' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function





' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Private Sub Text_LostFocus(Index As Integer)

    Select Case Index
        Case ptxSOKO
            Text(Index).Text = StrConv(Trim(Text(Index).Text), vbUpperCase)
    End Select


End Sub
