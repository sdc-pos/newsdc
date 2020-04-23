VERSION 5.00
Begin VB.Form F1070101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "棚卸し表印刷"
   ClientHeight    =   6945
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
   ScaleHeight     =   6945
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   5640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   8955
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   8355
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   7755
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6435
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   5835
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   5835
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "処理選択"
      Height          =   2175
      Left            =   480
      TabIndex        =   23
      Top             =   840
      Width           =   2775
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "印刷範囲クリアー"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "再印刷"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "新規印刷"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   1
      Left            =   960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実　行"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "倉庫"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   ")"
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   34
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblZEN_LOC 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   5760
      TabIndex        =   33
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "（前回指示範囲"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   32
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   4680
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   8715
      TabIndex        =   29
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   8115
      TabIndex        =   28
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "～"
      Height          =   255
      Index           =   4
      Left            =   7515
      TabIndex        =   27
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   6795
      TabIndex        =   26
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   6195
      TabIndex        =   25
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "標準棚番範囲"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   24
      Top             =   2280
      Width           =   1575
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
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1070101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSOKO% = 0                  '対象倉庫
Private Const ptxS_RETU% = 1                '開始　標準棚番　列
Private Const ptxS_REN% = 2                 '開始　標準棚番　連
Private Const ptxS_DAN% = 3                 '開始　標準棚番　段
Private Const ptxE_RETU% = 4                '終了　標準棚番　列
Private Const ptxE_REN% = 5                 '終了　標準棚番　連
Private Const ptxE_DAN% = 6                 '終了　標準棚番　段

Private Const Text_Max% = 6                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNaigai% = 0               '国内外

                                            '別置きロケーション検索用
Private Type Betu_Loc_Tag
    SOKO        As String * 2
    Retu        As String * 2
    Ren         As String * 2
    Dan         As String * 2
    ZAIKO_QTY   As Long
End Type

Private Betu_Loc(0 To 2)    As Betu_Loc_Tag

Private Const LMAX% = 41                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）


Dim NormalFont As New StdFont               '印刷フォント


'Private Const Last_Update_Day$ = "(F107010) 2018.04.11 14:45"
'Private Const Last_Update_Day$ = "(F107010) 2018.11.16 13:00"
Private Const Last_Update_Day$ = "(F107010) 2020.01.16 17:00 物流 差異項目追加"


Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   エラーチェック処理
'----------------------------------------------------------------------------
                                            
Dim i       As Integer
Dim sts     As Integer

Dim ans     As Integer
                                            
Dim c       As String * 128
                                            
                                            
Dim NEXT_F  As Integer
                                            
    Err_Chk = True

    
    
    For i = ptxSOKO To ptxE_DAN
        If IsNumeric(Text(i).Text) Then
            Text(i).Text = Format(Text(i).Text, "00")
        End If
    Next i
    
    
    
    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            
            If GetIni(App.EXEName, "ZENKAI_LOC" & Trim(Text(ptxSOKO).Text), "SYS", c) Then
                lblZEN_LOC.Caption = ""
            Else
                lblZEN_LOC.Caption = RTrim(c)
            End If
        
        Case BtErrKeyNotFound
        
            Beep
            MsgBox ("倉庫未登録です。")
            Text(ptxSOKO).SetFocus
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetGreaterEqual, "倉庫ﾏｽﾀ")
            Exit Function
    
    End Select
    
    
    
    
    
    If (Text(ptxS_RETU).Text & Text(ptxS_REN).Text & Text(ptxS_DAN).Text) > _
        (Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
        Beep
        MsgBox ("入力した項目はエラーです。")
        Text(ptxSOKO).SetFocus
        Exit Function
    End If
    
    If Option1(0).Value Then
                                
        '新規印刷処理
        If Last_JGYOBU = "*" Then
        '---------------    全BU
                    
            Call UniCode_Conv(K5_STOCK.NAIGAI, Right(Combo(pcmbNaigai), 1))
            Call UniCode_Conv(K5_STOCK.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K5_STOCK.CHECK_MARK, "")
            sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K5_STOCK, Len(K5_STOCK), 5)
            Select Case sts
                Case BtNoErr
                        
                    
                    If StrConv(STOCKREC.JGYOBU, vbUnicode) <> SHIZAI And _
                     StrConv(STOCKREC.ST_SOKO, vbUnicode) = Text(ptxSOKO).Text Then
                        If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Then
                            Beep
                            ans = MsgBox("未処理の棚卸しデータが有ります。このまま処理を継続しますか？", vbYesNo + vbDefaultButton1)
                            If ans = vbYes Then
                                
                            Else
                                Exit Function
                            End If
                        End If
                        
                    End If
                Case BtErrEOF
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
                
            Call UniCode_Conv(K4_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K4_STOCK.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K4_STOCK.ST_RETU, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K4_STOCK.ST_REN, Text(ptxS_REN).Text)
            Call UniCode_Conv(K4_STOCK.ST_DAN, Text(ptxS_DAN).Text)
            Call UniCode_Conv(K4_STOCK.HIN_GAI, "")
            
            sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K4_STOCK, Len(K4_STOCK), 4)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(STOCKREC.NAIGAI, vbUnicode) = Right(Combo(pcmbNaigai).Text, 1) Then
                        If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                            <= (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Beep
                            ans = MsgBox("指定された範囲内に処理中データが有ります。このまま処理を継続しますか？（既存の情報はクリアーされます）", vbYesNo + vbDefaultButton1)
                            If ans = vbYes Then
                            Else
                                Exit Function
                            End If
                        End If
                    End If
                Case BtErrEOF
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
        Else
        '---------------    個別BU
            Call UniCode_Conv(K2_STOCK.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_STOCK.NAIGAI, Right(Combo(pcmbNaigai), 1))
            Call UniCode_Conv(K2_STOCK.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K2_STOCK.CHECK_MARK, "")
            sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K2_STOCK, Len(K2_STOCK), 2)
            Select Case sts
                Case BtNoErr
                    If StrConv(STOCKREC.ST_SOKO, vbUnicode) = Text(ptxSOKO).Text Then
                        If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Then
                            Beep
                            ans = MsgBox("未処理の棚卸しデータが有ります。このまま処理を継続しますか？", vbYesNo + vbDefaultButton1)
                            If ans = vbYes Then
                            Else
                                Exit Function
                            End If
                        End If
                    End If
                Case BtErrEOF
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
        
            Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
            Call UniCode_Conv(K1_STOCK.ST_RETU, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K1_STOCK.ST_REN, Text(ptxS_REN).Text)
            Call UniCode_Conv(K1_STOCK.ST_DAN, Text(ptxS_DAN).Text)
            Call UniCode_Conv(K1_STOCK.HIN_GAI, "")
            
            sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
            Select Case sts
                Case BtNoErr
                    If StrConv(STOCKREC.JGYOBU, vbUnicode) = Last_JGYOBU And _
                        StrConv(STOCKREC.NAIGAI, vbUnicode) = Right(Combo(pcmbNaigai).Text, 1) Then
                        If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                            <= (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Beep
                            ans = MsgBox("指定された範囲内に処理中データが有ります。このまま処理を継続しますか？（既存の情報はクリアーされます）", vbYesNo + vbDefaultButton1)
                            If ans = vbYes Then
                            Else
                                Exit Function
                            End If
                        End If
                    End If
                Case BtErrEOF
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
        End If
    End If
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1070101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070101)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070101)

    F1070101.MousePointer = vbDefault

End Sub

Private Sub Command_Click(Index As Integer)

Dim yn          As Integer
Dim i           As Integer
Dim mesg        As String
Dim Data_cnt    As Long     '2018.04.11
    
    Select Case Index
        Case 0                              '実行
            If Err_Chk() Then
                Exit Sub
            End If
            
            If Option1(0).Value Then
                mesg = "新規棚卸し表印刷"
            
                Data_cnt = New_Count_Proc()
            
            End If
            
            If Option1(1).Value Then
                
                mesg = "棚卸し表再印刷"
            
                Data_cnt = Count_Proc()
            
            End If
            
            If Option1(2).Value Then
                mesg = "棚卸し情報クリア―処理"
            End If
            
            If Option1(0).Value Or Option1(1).Value Then
                mesg = mesg & Chr(13) & Chr(10)
                mesg = mesg & "印刷件数= " & Data_cnt & "件" & Chr(13) & Chr(10)
                mesg = mesg & "実行しますか？"
                yn = MsgBox("「棚卸し」" & mesg, vbYesNo + vbQuestion, "確認入力")
            Else
                yn = MsgBox("「棚卸し」" & mesg & "実行しますか？", vbYesNo + vbQuestion, "確認入力")
            End If

            If yn = vbYes Then
                If Option1(0).Value Then        '新規印刷
                    If New_Print_Proc() Then
                        Unload Me
                    End If
                End If
                If Option1(1).Value Then        '再印刷
                    If Print_Proc() Then
                        Unload Me
                    End If
                End If
                If Option1(2).Value Then        'データクリアー
                    If Data_Clear_Proc() Then
                        Unload Me
                    End If
                End If
            End If
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
End Sub

Private Sub Form_DblClick()
'    PrintForm                  '2018.04.09
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

Dim c       As String * 128
Dim i       As Integer

Dim varWK   As Variant      '2008.07.24
     
'     If App.PrevInstance Then                  2018.04.09
'        Beep                                   2018.04.09
'        MsgBox "同一プログラム実行中です。"    2018.04.09
'        End                                    2018.04.09
'    End If                                     2018.04.09
    
                                'ログファイル名取り込み
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
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
    
    '全BU可／不可
    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = ""
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "-"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12

    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "全BU"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12

    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = ""
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "-"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12

    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If

        Load SubMenu(i + 1)
        
        If RTrim(JGYOBU_T(i).NAME) = "-" Then
            SubMenu(i).Checked = False
        End If
        
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If Trim(JGYOBU_T(i).CODE) = "" Then
        Else
            If JGYOBU_T(i).CODE = Last_JGYOBU Then
                F1070101.Caption = "棚卸し表印刷（" & RTrim(JGYOBU_T(i).NAME) & ") " & Last_Update_Day
                SubMenu(i).Checked = True
                LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
    '            LabJIGYO.BorderStyle = 1
            Else
                SubMenu(i).Checked = False
            End If
        End If
    Next i
    Unload SubMenu(i)
                                
''                                '前回指示範囲（新規指示） 2007.08.22
''    If GetIni(App.EXEName, "ZENKAI_LOC", "SYS", c) Then   2007.08.22
''        lblZEN_LOC.Caption = ""                           2007.08.22
''    Else                                                  2007.08.22
''        lblZEN_LOC.Caption = RTrim(c)                     2007.08.22
''    End If                                                2007.08.22
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
                                '棚卸しデータファイルＯＰＥＮ
    If STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫集計データファイルＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1070101.FontName
        .Size = 11
    End With

    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    Show

    Option1(0).Value = True
    Text(ptxSOKO).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_ITEM), 0)
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
                                            '棚卸しデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚卸しデータ")
        End If
    End If
                                            '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070101 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
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
    F1070101.Caption = "棚卸し表印刷（" & RTrim(JGYOBU_T(i).NAME) & ") " & Last_Update_Day
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
    
Dim c   As String * 128
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxSOKO
            Text(Index).Text = Trim(StrConv(Text(Index).Text, vbUpperCase))
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    If GetIni(App.EXEName, "ZENKAI_LOC" & Trim(Text(ptxSOKO).Text), App.EXEName, c) Then
                        lblZEN_LOC.Caption = ""
                    Else
                        lblZEN_LOC.Caption = RTrim(c)
                    End If
                Case BtErrKeyNotFound
                    Beep
                    MsgBox ("倉庫未登録です。")
                    Text(ptxSOKO).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "倉庫ﾏｽﾀ")
                    Exit Sub
            End Select
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub

Private Function New_Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   新規印刷処理
'----------------------------------------------------------------------------
Dim com             As Integer
Dim sts             As Integer
Dim Sumi_Zaiko_Qty  As Long
Dim Mi_Zaiko_Qty    As Long
Dim i               As Integer
Dim j               As Integer
Dim POS_Zaiko_Qty   As Long
Dim Betu_Location   As String * 8
Dim Betu_Zaiko_Qty  As Long
Dim ans             As Integer
Dim Print_cnt       As Integer      '2018.04.09

    New_Print_Proc = True
    If Data_Clear_Proc() Then       'データ初期化
        Exit Function
    End If

    Call Input_Lock
    
    If Last_JGYOBU = "*" Then
        '全BU
        For i = 0 To UBound(JGYOBU_T)
        
            If JGYOBU_T(i).CODE = "*" Or JGYOBU_T(i).CODE = SHIZAI Then
            Else
        
                Call UniCode_Conv(K6_ITEM.JGYOBU, JGYOBU_T(i).CODE)
                Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxSOKO).Text)
                Call UniCode_Conv(K6_ITEM.ST_RETU, Text(ptxS_RETU).Text)
                Call UniCode_Conv(K6_ITEM.ST_REN, Text(ptxS_REN).Text)
                Call UniCode_Conv(K6_ITEM.ST_DAN, Text(ptxS_DAN).Text)
                Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
                
                com = BtOpGetGreaterEqual
            
                Do
                    
                    sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
                    Select Case sts
                        Case BtNoErr
                            
                            If StrConv(ITEMREC.JGYOBU, vbUnicode) <> JGYOBU_T(i).CODE Or _
                                StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                                Exit Do
                            End If
                                    
                            If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) > _
                                (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            
                            Exit Do
                        
                        Case Else
                            Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                            Exit Function
                    End Select
                                                        '該当品番が棚卸データに存在した場合は削除する（標準棚番変更時の対応）
                    Call UniCode_Conv(K0_STOCK.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_STOCK.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_STOCK.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        
                    Do
                        sts = BTRV(BtOpGetEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K1_STOCK), 0)
                        Select Case sts
                            Case BtNoErr

                                Do
                                    sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                            Beep
                                            ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                            If ans = vbCancel Then
                                                Exit Function
                                            End If
                                        Case Else
                                            Call File_Error(sts, BtOpDelete, "棚卸しデータ")
                                            Exit Function
                                    End Select
                                Loop
                        
                                Exit Do
                            Case BtErrKeyNotFound
                                Exit Do
                            
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "棚卸しデータ")
                                Exit Function
                        End Select
                    Loop
                                                    '棚卸しデータ作成
                    Call UniCode_Conv(STOCKREC.JGYOBU, JGYOBU_T(i).CODE)                        '事業部
                    Call UniCode_Conv(STOCKREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))        '国内外
                    Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '品目コード
                    Call UniCode_Conv(STOCKREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))    '標準棚番
                    Call UniCode_Conv(STOCKREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(STOCKREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(STOCKREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))       '松下理論在庫
                    Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                            Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                            Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                            Exit Function
                    End Select
                    Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                    Call UniCode_Conv(STOCKREC.BU_ZAI_QTY, StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode))
                    Call UniCode_Conv(STOCKREC.PPSC_ZAI_QTY, StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode))
                                                                                            'POS総在庫集計
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        
                        Exit Function
                    End If
                    Call UniCode_Conv(STOCKREC.POS_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                    POS_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                                            
                    For j = 0 To UBound(Betu_Loc)
                        Betu_Loc(j).SOKO = ""
                        Betu_Loc(j).Retu = ""
                        Betu_Loc(j).Ren = ""
                        Betu_Loc(j).Dan = ""
                        Betu_Loc(j).ZAIKO_QTY = 0
                    Next j
                    Betu_Zaiko_Qty = 0
                                                                                            '標準棚番在庫集計
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                            StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                        Exit Function
                    End If
                    
                    Call UniCode_Conv(STOCKREC.ST_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                    Call UniCode_Conv(STOCKREC.EE1_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE1_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.EE2_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE2_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.ETC_ZAIKO, "00000000")
                    
                    Betu_Loc(0).SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Betu_Loc(0).Retu = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Betu_Loc(0).Ren = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Betu_Loc(0).Dan = StrConv(ITEMREC.ST_DAN, vbUnicode)
                    Betu_Loc(0).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                    
                    Betu_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                        
                    If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                            '別置き検索
                    Else
                        For j = 1 To UBound(Betu_Loc)
                        
                            If Tana_Kensaku(Betu_Location) Then
                                Exit Function
                            End If
                    
                            If Len(Trim(Betu_Location)) = 0 Then
                                                            '検索完了
                                Exit For
                            End If
                                
                            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                            Betu_Location) Then
                                Exit Function
                            End If
                        
                            Betu_Loc(j).SOKO = Left(Betu_Location, 2)
                            Betu_Loc(j).Retu = Mid(Betu_Location, 3, 2)
                            Betu_Loc(j).Ren = Mid(Betu_Location, 5, 2)
                            Betu_Loc(j).Dan = Right(Betu_Location, 2)
                            Betu_Loc(j).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                            
                            Betu_Zaiko_Qty = Betu_Zaiko_Qty + (Sumi_Zaiko_Qty + Mi_Zaiko_Qty)
                    
                    
                            If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                        '検索完了
                                Exit For
                            End If
                    
                        Next j
                                                        '別置き１
                        If Betu_Loc(1).ZAIKO_QTY <> 0 Then
                            Call UniCode_Conv(STOCKREC.EE1_LOCATION, (Betu_Loc(1).SOKO & Betu_Loc(1).Retu & Betu_Loc(1).Ren & Betu_Loc(1).Dan))
                            Call UniCode_Conv(STOCKREC.EE1_ZAIKO, Format(Betu_Loc(1).ZAIKO_QTY, "00000000"))
                        End If
                                                        '別置き２
                        If Betu_Loc(2).ZAIKO_QTY <> 0 Then
                            Call UniCode_Conv(STOCKREC.EE2_LOCATION, (Betu_Loc(2).SOKO & Betu_Loc(2).Retu & Betu_Loc(2).Ren & Betu_Loc(2).Dan))
                            Call UniCode_Conv(STOCKREC.EE2_ZAIKO, Format(Betu_Loc(2).ZAIKO_QTY, "00000000"))
                        End If
                                                        '別置き３
                        Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                        Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                                                        '別置き４
                        Call UniCode_Conv(STOCKREC.ETC_ZAIKO, Format((POS_Zaiko_Qty - Betu_Zaiko_Qty), "00000000"))

                    End If
                    
                    Call UniCode_Conv(STOCKREC.CHECK_MARK, "")                      '照合マーク
                    Call UniCode_Conv(STOCKREC.PRINT_YMD, Format(Now, "YYYYMMDD"))  '印刷日付
                    Call UniCode_Conv(STOCKREC.INPUT_YMD, "")                       '入力日付
                    Call UniCode_Conv(STOCKREC.SAI_QTY, "000000000")                '差異数
                    Call UniCode_Conv(STOCKREC.FILLER, "")

                    If CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)) = 0 And _
                        CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)) = 0 Then
                    Else
                        Do
                            sts = BTRV(BtOpInsert, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case Else
                                    Call File_Error(sts, BtOpInsert, "棚卸しデータ")
                                    Exit Function
                            End Select
                    
                        Loop
                    End If
                    com = BtOpGetNext
                Loop
            End If
        Next i
    Else
  
        '単独BU
        Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K6_ITEM.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K6_ITEM.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K6_ITEM.ST_DAN, Text(ptxS_DAN).Text)
        Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
        
        com = BtOpGetGreaterEqual

        Do
  
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                            
                    If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) > _
                        (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    
                    Exit Do
                
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
                                                '該当品番が棚卸データに存在した場合は削除する（標準棚番変更時の対応）
            Call UniCode_Conv(K0_STOCK.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_STOCK.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_STOCK.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
                
            Do
    '            sts = BTRV(BtOpGetEqual + BtSNoWait, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K1_STOCK), 0)
                sts = BTRV(BtOpGetEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K1_STOCK), 0)
                Select Case sts
                    Case BtNoErr
                
                
                        Do
                            sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "棚卸しデータ")
                                    Exit Function
                            End Select
                        Loop
                
                
                        Exit Do
                    Case BtErrKeyNotFound
                        Exit Do
                    
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "棚卸しデータ")
                        Exit Function
                End Select
            Loop
                                            '棚卸しデータ作成
            Call UniCode_Conv(STOCKREC.JGYOBU, Last_JGYOBU)                             '事業部
            Call UniCode_Conv(STOCKREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))        '国内外
            Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '品目コード
                                                                                        '標準棚番
            Call UniCode_Conv(STOCKREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                                                        '松下理論在庫
            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                
                    Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                    Exit Function
            End Select
            Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                                                                                    
            Call UniCode_Conv(STOCKREC.BU_ZAI_QTY, StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode))
            Call UniCode_Conv(STOCKREC.PPSC_ZAI_QTY, StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode))
                                                                                    
                                                                                    'ＰＯＳ総在庫集計
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                Exit Function
            End If
            Call UniCode_Conv(STOCKREC.POS_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
            POS_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                                    
            For i = 0 To UBound(Betu_Loc)
                Betu_Loc(i).SOKO = ""
                Betu_Loc(i).Retu = ""
                Betu_Loc(i).Ren = ""
                Betu_Loc(i).Dan = ""
                Betu_Loc(i).ZAIKO_QTY = 0
            Next i
            Betu_Zaiko_Qty = 0
                                                                                    '標準棚番在庫集計
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                    StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                
                Exit Function
            End If
            
            Call UniCode_Conv(STOCKREC.ST_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
            
            Call UniCode_Conv(STOCKREC.EE1_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE1_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.EE2_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE2_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.ETC_ZAIKO, "00000000")
            
            
            Betu_Loc(0).SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Betu_Loc(0).Retu = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Betu_Loc(0).Ren = StrConv(ITEMREC.ST_REN, vbUnicode)
            Betu_Loc(0).Dan = StrConv(ITEMREC.ST_DAN, vbUnicode)
            Betu_Loc(0).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
            
            Betu_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                
            If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                    '別置き検索
            Else
                For i = 1 To UBound(Betu_Loc)
                
                    If Tana_Kensaku(Betu_Location) Then
                        Exit Function
                    End If
            
                    If Len(Trim(Betu_Location)) = 0 Then
                                                    '検索完了
                        Exit For
                    End If
                        
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                    Betu_Location) Then
                        Exit Function
                    End If
                
                    Betu_Loc(i).SOKO = Left(Betu_Location, 2)
                    Betu_Loc(i).Retu = Mid(Betu_Location, 3, 2)
                    Betu_Loc(i).Ren = Mid(Betu_Location, 5, 2)
                    Betu_Loc(i).Dan = Right(Betu_Location, 2)
                    Betu_Loc(i).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                    
                    Betu_Zaiko_Qty = Betu_Zaiko_Qty + (Sumi_Zaiko_Qty + Mi_Zaiko_Qty)
            
            
                    If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                '検索完了
                        Exit For
                    End If
            
                Next i
            
            
                                                '別置き１
                If Betu_Loc(1).ZAIKO_QTY <> 0 Then
                    Call UniCode_Conv(STOCKREC.EE1_LOCATION, (Betu_Loc(1).SOKO & Betu_Loc(1).Retu & Betu_Loc(1).Ren & Betu_Loc(1).Dan))
                    Call UniCode_Conv(STOCKREC.EE1_ZAIKO, Format(Betu_Loc(1).ZAIKO_QTY, "00000000"))
                End If
                                                '別置き２
                If Betu_Loc(2).ZAIKO_QTY <> 0 Then
                    Call UniCode_Conv(STOCKREC.EE2_LOCATION, (Betu_Loc(2).SOKO & Betu_Loc(2).Retu & Betu_Loc(2).Ren & Betu_Loc(2).Dan))
                    Call UniCode_Conv(STOCKREC.EE2_ZAIKO, Format(Betu_Loc(2).ZAIKO_QTY, "00000000"))
                End If
                                                '別置き３
                Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                                                '別置き４
                Call UniCode_Conv(STOCKREC.ETC_ZAIKO, Format((POS_Zaiko_Qty - Betu_Zaiko_Qty), "00000000"))
            
            
            End If
            
            Call UniCode_Conv(STOCKREC.CHECK_MARK, "")                      '照合マーク
            Call UniCode_Conv(STOCKREC.PRINT_YMD, Format(Now, "YYYYMMDD"))  '印刷日付
            Call UniCode_Conv(STOCKREC.INPUT_YMD, "")                       '入力日付
        
            Call UniCode_Conv(STOCKREC.SAI_QTY, "000000000")                '差異数
        
        
            Call UniCode_Conv(STOCKREC.FILLER, "")
        
        
        
            If CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)) = 0 And _
                CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)) = 0 Then
            Else
                Do
                    sts = BTRV(BtOpInsert, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        
                        Case Else
                            Call File_Error(sts, BtOpInsert, "棚卸しデータ")
                            Exit Function
                    End Select
            
                Loop
            End If
        
            com = BtOpGetNext
        Loop
    End If

    Call Input_UnLock
    
    If Print_Proc(Print_cnt) Then
        Exit Function
    End If

    lblZEN_LOC.Caption = Text(ptxSOKO).Text & "-" & Text(ptxS_RETU).Text & "-" & Text(ptxS_REN).Text & "-" & Text(ptxS_DAN).Text & "～" _
                            & Text(ptxSOKO).Text & "-" & Text(ptxE_RETU).Text & "-" & Text(ptxE_REN).Text & "-" & Text(ptxE_DAN).Text & "   " & Format(Print_cnt) & "件"
    
    If WriteIni(App.EXEName, "ZENKAI_LOC" & Trim(Text(ptxSOKO).Text), App.EXEName, Trim(lblZEN_LOC.Caption)) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & " ZENKAI_LOC")
        Exit Function
    End If

    New_Print_Proc = False

End Function
Private Function Data_Clear_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ初期化処理
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer
Dim i       As Integer

    Data_Clear_Proc = True

    Call Input_Lock
    
    If Last_JGYOBU = "*" Then
        '全BU
        For i = 0 To UBound(JGYOBU_T)
            If JGYOBU_T(i).CODE = "*" Or JGYOBU_T(i).CODE = SHIZAI Then
            Else
                Call UniCode_Conv(K1_STOCK.JGYOBU, JGYOBU_T(i).CODE)
                Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
                Call UniCode_Conv(K1_STOCK.ST_RETU, Text(ptxS_RETU).Text)
                Call UniCode_Conv(K1_STOCK.ST_REN, Text(ptxS_REN).Text)
                Call UniCode_Conv(K1_STOCK.ST_DAN, Text(ptxS_DAN).Text)
                Call UniCode_Conv(K1_STOCK.HIN_GAI, "")
            
                com = BtOpGetGreater
            
                Do
                    DoEvents
                    Do
                        sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                        Select Case sts
                            Case BtNoErr
                                If StrConv(STOCKREC.JGYOBU, vbUnicode) <> JGYOBU_T(i).CODE Or _
                                    StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                                    sts = BtErrEOF
                                End If
                                If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                                    > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                                    sts = BtErrEOF
                                End If
                                Exit Do
                            Case BtErrEOF
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                                Exit Function
                        End Select
                    Loop
                    If sts = BtErrEOF Then
                        Exit Do
                    End If
                    Do
                        sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "棚卸しデータ")
                                Exit Function

                        End Select
                    Loop
                    com = BtOpGetNext
                Loop
            End If
        Next i

    Else
        '単独BU
        Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K1_STOCK.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K1_STOCK.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K1_STOCK.ST_DAN, Text(ptxS_DAN).Text)
        Call UniCode_Conv(K1_STOCK.HIN_GAI, "")

        com = BtOpGetGreater

        Do
            DoEvents
            Do
                Dim iniSUMZREC As SUMZREC_Tag '2020/04/22
                SUMZREC = iniSUMZREC
            
                sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                Select Case sts
                    Case BtNoErr
                        If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                            StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                            sts = BtErrEOF
                        End If
                        If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                            > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            sts = BtErrEOF
                        End If
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                        Exit Function
                End Select
            Loop
            If sts = BtErrEOF Then
                Exit Do
            End If
            Do
                sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "棚卸しデータ")
                        Exit Function
                End Select
            Loop
            com = BtOpGetNext
        Loop
    End If
    
    Call Input_UnLock
    
    Data_Clear_Proc = False
End Function
Private Function Tana_Kensaku(Betu_Location As String) As Integer
'----------------------------------------------------------------------------
'                   別置き検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim i           As Integer
    
Dim Check_Flg   As Integer

    Tana_Kensaku = True
    
    Betu_Location = ""
    
    
    Call UniCode_Conv(K4_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K4_ZAIKO.Retu, "")
    Call UniCode_Conv(K4_ZAIKO.Ren, "")
    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        Select Case sts
            Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                    Exit Do
                Else
                    Check_Flg = 0
                    For i = 0 To UBound(Betu_Loc)
                        If Len(Trim(Betu_Loc(i).SOKO)) = 0 Then
                            Exit For
                        End If
                
                        If Betu_Loc(i).SOKO = StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
                            Betu_Loc(i).Retu = StrConv(ZAIKOREC.Retu, vbUnicode) And _
                            Betu_Loc(i).Ren = StrConv(ZAIKOREC.Ren, vbUnicode) And _
                            Betu_Loc(i).Dan = StrConv(ZAIKOREC.Dan, vbUnicode) Then
                            Check_Flg = 1
                            Exit For
                        End If
                    Next i
                                
                
                    If Check_Flg = 0 Then
                        Betu_Location = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                        Exit Do
                    End If
            
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
            
    Loop
    
    Tana_Kensaku = False

End Function
Private Function Print_Proc(Optional Print_cnt As Integer) As Integer
'----------------------------------------------------------------------------
'                   棚卸し表印刷処理
'----------------------------------------------------------------------------
Dim Lcnt        As Integer

Dim sts         As Integer
Dim com         As Integer

Dim Save_Soko   As String * 2

Dim Edit        As String

Dim X_Tab       As Integer

'Dim Print_cnt   As Integer  '2007.12.03

    Print_Proc = True

    Call Input_Lock


    Lcnt = LMAX

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    Print_cnt = 0           '2007.12.03
    
    
    
    
    If Last_JGYOBU = "*" Then
        '全BU
    
        Call UniCode_Conv(K4_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        
        Call UniCode_Conv(K4_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K4_STOCK.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K4_STOCK.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K4_STOCK.ST_DAN, Text(ptxS_DAN).Text)
        
        Call UniCode_Conv(K4_STOCK.HIN_GAI, "")
        
        com = BtOpGetGreaterEqual
        
        
        Do
            sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K4_STOCK, Len(K4_STOCK), 4)
            Select Case sts
                Case BtNoErr
                     If StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                                            
                    If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                        > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                    
                Case BtErrEOF
                    Exit Do
                    
                    
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                    Exit Function
            End Select
    '-------------------------------------------------  明細印刷
            If com = BtOpGetGreaterEqual Then
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
            End If
            
            If Save_Soko <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
                
            End If
            
            
            If Head_Print_Proc(Lcnt, Save_Soko) Then
                Exit Function
            End If
                
            X_Tab = MGN_L
                
            Printer.Print Tab(X_Tab);
                
            Printer.Print Left(StrConv(STOCKREC.HIN_GAI, vbUnicode), 14);    '品番（外部）      '2015.12.24
            X_Tab = X_Tab + 15
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '理論在庫
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
            
            If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
                Edit = Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0")
            Else
                Edit = "0"
            End If
            
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     'Active在庫
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                        
            If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
                Edit = Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0")
            Else
                Edit = "0"
            End If
            
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     'GLICS在庫
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
            
            Edit = Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     'POS在庫
                            
            X_Tab = X_Tab + Len(Edit) + 2
            Printer.Print Tab(X_Tab);
            
            Edit = Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0") - Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '差異 2020/01/16追加

            X_Tab = X_Tab + Len(Edit) + 5
            Printer.Print Tab(X_Tab);
            
            Edit = StrConv(STOCKREC.ST_RETU, vbUnicode) & "-" & _
                    StrConv(STOCKREC.ST_REN, vbUnicode) & "-" & _
                    StrConv(STOCKREC.ST_DAN, vbUnicode)
            Printer.Print Edit;                                     '標準棚番
                
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
            
            Edit = Format(CLng(StrConv(STOCKREC.ST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '標準棚番在庫
            
            '------------------------------------------------------ '別置き在庫１▼
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            If CLng(StrConv(STOCKREC.EE1_ZAIKO, vbUnicode)) = 0 Then
                Edit = Space(11)
            Else
                Edit = Left(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 3, 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 5, 2) & "-" & _
                        Right(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 2)
            End If
            Printer.Print Edit;
                
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.EE1_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------ '別置き在庫１▲
            '------------------------------------------------------ '別置き在庫２▼
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            If CLng(StrConv(STOCKREC.EE2_ZAIKO, vbUnicode)) = 0 Then
                Edit = Space(11)
            Else
                Edit = Left(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 3, 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 5, 2) & "-" & _
                        Right(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 2)
            End If
            Printer.Print Edit;
                
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.EE2_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------ '別置き在庫２▲
            '------------------------------------------------------ 'その他在庫▼
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.ETC_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------ 'その他在庫▲
    
            X_Tab = X_Tab + Len(Edit)
            Printer.Print Tab(X_Tab);
                
'            Printer.Print "(           )"  '2020/01/16 括弧削除 2020/04/14
            Edit = "(" & Format(CLng(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)), "#") & ")"
            Printer.Print Edit;
                        
            Printer.Print String(130, "─")
            Lcnt = Lcnt + 2
            
            Print_cnt = Print_cnt + 1       '2007.12.03
            
            com = BtOpGetNext
        Loop
    Else
        '単独BU
        Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K1_STOCK.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K1_STOCK.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K1_STOCK.ST_DAN, Text(ptxS_DAN).Text)
        Call UniCode_Conv(K1_STOCK.HIN_GAI, "")

        com = BtOpGetGreaterEqual
        
        Do
            sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
            Select Case sts
                Case BtNoErr
                    If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                                            
                    If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                        > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                    
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                    Exit Function
            End Select
    '-------------------------------------------------  明細印刷
            If com = BtOpGetGreaterEqual Then
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
            End If
            
            If Save_Soko <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
                
            End If
            
            
            If Head_Print_Proc(Lcnt, Save_Soko) Then
                Exit Function
            End If
                
            X_Tab = MGN_L
                
            Printer.Print Tab(X_Tab);

            Printer.Print Left(StrConv(STOCKREC.HIN_GAI, vbUnicode), 14);    '品番（外部）  2015.12.24
            X_Tab = X_Tab + 15
            Printer.Print Tab(X_Tab);
                
                
            Edit = Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '理論在庫
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                        
                        
            Edit = Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     'PPSC在庫
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
            
            '------------------------------------------------------GLICS在庫▼
            Edit = Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
            '------------------------------------------------------GLICS在庫▲
            
            
                        
            Edit = Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     'POS在庫
                            
            X_Tab = X_Tab + Len(Edit) + 2
            Printer.Print Tab(X_Tab);
                        
            
                        
            Edit = Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0") - Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '差異項目追加 2020/01/16

            X_Tab = X_Tab + Len(Edit) + 5
            Printer.Print Tab(X_Tab);
            
            
            
                            
            Edit = StrConv(STOCKREC.ST_RETU, vbUnicode) & "-" & _
                    StrConv(STOCKREC.ST_REN, vbUnicode) & "-" & _
                    StrConv(STOCKREC.ST_DAN, vbUnicode)
            Printer.Print Edit;                                     '標準棚番
                
            X_Tab = X_Tab + Len(Edit) + 2
            Printer.Print Tab(X_Tab);
    
            Edit = Format(CLng(StrConv(STOCKREC.ST_ZAIKO, vbUnicode)), "#0")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;                                     '標準棚番在庫
                
    
                
            '------------------------------------------------------別置き在庫１▽
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            If CLng(StrConv(STOCKREC.EE1_ZAIKO, vbUnicode)) = 0 Then
                Edit = Space(11)
            Else
                Edit = Left(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 3, 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 5, 2) & "-" & _
                        Right(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 2)
            End If
            Printer.Print Edit;
                
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.EE1_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------別置き在庫１△
                
            '------------------------------------------------------別置き在庫２▽
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            If CLng(StrConv(STOCKREC.EE2_ZAIKO, vbUnicode)) = 0 Then
                Edit = Space(11)
            Else
                Edit = Left(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 3, 2) & "-" & _
                        Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 5, 2) & "-" & _
                        Right(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 2)
            End If
            Printer.Print Edit;
                
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.EE2_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------別置在庫2▲
    
    
            '------------------------------------------------------その他在庫▼
            X_Tab = X_Tab + Len(Edit) + 3
            Printer.Print Tab(X_Tab);
                
            Edit = Format(CLng(StrConv(STOCKREC.ETC_ZAIKO, vbUnicode)), "#")
            If Len(Edit) < 6 Then
                Edit = Space(6 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
            '------------------------------------------------------その他在庫▲
    
            X_Tab = X_Tab + Len(Edit)
            Printer.Print Tab(X_Tab);
    
    '        Printer.Print "(           )"  '2020/01/16 括弧コメントアウト 2020/04/20

            Printer.Print ""; '2020/04/20 セミコロンを付けると下のprinter.printと同じ行になる
            Edit = "(" & Right(Space(11) & Format(CLng(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)), "#"), 11) & ")"
            Printer.Print Edit
                                    
            Printer.Print String(130, "─")
            Lcnt = Lcnt + 2
            
            Print_cnt = Print_cnt + 1       '2007.12.03
            
            
            com = BtOpGetNext
        
        Loop


    End If

    Printer.EndDoc


    MsgBox "「" & StrConv(Format(Print_cnt, "#,##0"), vbWide) & "」件の印刷を行いました。"  '2007.12.03


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(Lcnt As Integer, Soko_No As String) As Integer
'----------------------------------------------------------------------------
'                   ヘッダーコントロール処理
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If Lcnt < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    If Lcnt = LMAX Then
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

    Printer.Print Tab(MGN_L + 60);
    
    Printer.Print "棚　卸　表";
    
    
    Printer.Print Tab(MGN_L + 100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫：";
    Call UniCode_Conv(K0_SOKO.Soko_No, Soko_No)
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode);
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
            Exit Function
    End Select
    Printer.Print

'---------------------2020/04/02 項目位置修正--------------------▼
    Printer.Print Tab(MGN_L);
    Printer.Print "品目コード";
    Printer.Print Tab(MGN_L + 16);
    Printer.Print "理 論";
    Printer.Print Tab(MGN_L + 24);
    Printer.Print "Active";
    Printer.Print Tab(MGN_L + 34);
    Printer.Print "Glics";
    Printer.Print Tab(MGN_L + 45);
    Printer.Print "POS";
    Printer.Print Tab(MGN_L + 52);
    Printer.Print "差異";                   '2020/01/15 差異項目追加
    Printer.Print Tab(MGN_L + 61);
    Printer.Print "標準棚番    在庫";
    Printer.Print Tab(MGN_L + 80);
    Printer.Print "別置1      別置1在庫";
    Printer.Print Tab(MGN_L + 102);
    Printer.Print "別置2        別置2在庫";
    Printer.Print Tab(MGN_L + 126);
    Printer.Print "その他"
'---------------------2020/04/02 項目位置修正--------------------▲
    Printer.Print
    
    Lcnt = 0
    
    Head_Print_Proc = False

End Function

Private Sub Text_LostFocus(Index As Integer)
    
    If Index = ptxSOKO Then
        Text(Index).Text = Trim(StrConv(Text(Index).Text, vbUpperCase))
    End If

End Sub
Private Function New_Count_Proc() As Long
'----------------------------------------------------------------------------
'                   新規印刷件数処理
'               2018.04.11
'----------------------------------------------------------------------------
Dim com             As Integer
Dim sts             As Integer

Dim Sumi_Zaiko_Qty  As Long
Dim Mi_Zaiko_Qty    As Long

Dim i               As Integer
Dim j               As Integer

Dim POS_Zaiko_Qty   As Long

Dim Betu_Location   As String * 8
Dim Betu_Zaiko_Qty  As Long

Dim ans             As Integer

Dim Data_cnt        As Long

    New_Count_Proc = True

    Data_cnt = 0

    Call Input_Lock
    
    
    
    If Last_JGYOBU = "*" Then
        '全BU
        For i = 0 To UBound(JGYOBU_T)
        
            If JGYOBU_T(i).CODE = "*" Or JGYOBU_T(i).CODE = SHIZAI Then
            Else
        
                Call UniCode_Conv(K6_ITEM.JGYOBU, JGYOBU_T(i).CODE)
                Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxSOKO).Text)
                Call UniCode_Conv(K6_ITEM.ST_RETU, Text(ptxS_RETU).Text)
                Call UniCode_Conv(K6_ITEM.ST_REN, Text(ptxS_REN).Text)
                Call UniCode_Conv(K6_ITEM.ST_DAN, Text(ptxS_DAN).Text)
                Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
                
                com = BtOpGetGreaterEqual
                
            
                Do
                    
                    sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
                    Select Case sts
                        Case BtNoErr
                            
                            If StrConv(ITEMREC.JGYOBU, vbUnicode) <> JGYOBU_T(i).CODE Or _
                                StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                                Exit Do
                            End If
                                    
                            If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) > _
                                (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            
                            Exit Do
                        
                        Case Else
                            Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                            Exit Function
                    End Select
                                                    '棚卸しデータ作成
                    Call UniCode_Conv(STOCKREC.JGYOBU, JGYOBU_T(i).CODE)                        '事業部
                    Call UniCode_Conv(STOCKREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))        '国内外
                    Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '品目コード
                                                                                                '標準棚番
                    Call UniCode_Conv(STOCKREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                    Call UniCode_Conv(STOCKREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(STOCKREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(STOCKREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                                                                '松下理論在庫
                    Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                            Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                            Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                            Exit Function
                    End Select
                    Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                    Call UniCode_Conv(STOCKREC.BU_ZAI_QTY, StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode))
                    Call UniCode_Conv(STOCKREC.PPSC_ZAI_QTY, StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode))
                                               
                                               'ＰＯＳ総在庫集計
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
                    Call UniCode_Conv(STOCKREC.POS_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                    POS_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                                            
                    For j = 0 To UBound(Betu_Loc)
                        Betu_Loc(j).SOKO = ""
                        Betu_Loc(j).Retu = ""
                        Betu_Loc(j).Ren = ""
                        Betu_Loc(j).Dan = ""
                        Betu_Loc(j).ZAIKO_QTY = 0
                    Next j
                    Betu_Zaiko_Qty = 0
                                               '標準棚番在庫集計
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                            StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                        Exit Function
                    End If
                    
                    Call UniCode_Conv(STOCKREC.ST_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                    Call UniCode_Conv(STOCKREC.EE1_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE1_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.EE2_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE2_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                    Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                    Call UniCode_Conv(STOCKREC.ETC_ZAIKO, "00000000")
                    
                    Betu_Loc(0).SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Betu_Loc(0).Retu = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Betu_Loc(0).Ren = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Betu_Loc(0).Dan = StrConv(ITEMREC.ST_DAN, vbUnicode)
                    Betu_Loc(0).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                    
                    Betu_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                        
                    If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                            '別置き検索
                    Else
                        For j = 1 To UBound(Betu_Loc)
                        
                            If Tana_Kensaku(Betu_Location) Then
                                Exit Function
                            End If
                    
                            If Len(Trim(Betu_Location)) = 0 Then
                                                            '検索完了
                                Exit For
                            End If
                                
                            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                            Mi_Zaiko_Qty, _
                                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                            Betu_Location) Then
                                Exit Function
                            End If
                        
                            Betu_Loc(j).SOKO = Left(Betu_Location, 2)
                            Betu_Loc(j).Retu = Mid(Betu_Location, 3, 2)
                            Betu_Loc(j).Ren = Mid(Betu_Location, 5, 2)
                            Betu_Loc(j).Dan = Right(Betu_Location, 2)
                            Betu_Loc(j).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                            
                            Betu_Zaiko_Qty = Betu_Zaiko_Qty + (Sumi_Zaiko_Qty + Mi_Zaiko_Qty)
                    
                    
                            If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                        '検索完了
                                Exit For
                            End If
                        Next j
                                                        '別置き１
                        If Betu_Loc(1).ZAIKO_QTY <> 0 Then
                            Call UniCode_Conv(STOCKREC.EE1_LOCATION, (Betu_Loc(1).SOKO & Betu_Loc(1).Retu & Betu_Loc(1).Ren & Betu_Loc(1).Dan))
                            Call UniCode_Conv(STOCKREC.EE1_ZAIKO, Format(Betu_Loc(1).ZAIKO_QTY, "00000000"))
                        End If
                                                        '別置き２
                        If Betu_Loc(2).ZAIKO_QTY <> 0 Then
                            Call UniCode_Conv(STOCKREC.EE2_LOCATION, (Betu_Loc(2).SOKO & Betu_Loc(2).Retu & Betu_Loc(2).Ren & Betu_Loc(2).Dan))
                            Call UniCode_Conv(STOCKREC.EE2_ZAIKO, Format(Betu_Loc(2).ZAIKO_QTY, "00000000"))
                        End If
                                                        '別置き３
                        Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                        Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                                                        '別置き４
                        Call UniCode_Conv(STOCKREC.ETC_ZAIKO, Format((POS_Zaiko_Qty - Betu_Zaiko_Qty), "00000000"))
                    End If
                    
                    Call UniCode_Conv(STOCKREC.CHECK_MARK, "")                      '照合マーク
                    Call UniCode_Conv(STOCKREC.PRINT_YMD, Format(Now, "YYYYMMDD"))  '印刷日付
                    Call UniCode_Conv(STOCKREC.INPUT_YMD, "")                       '入力日付
                    Call UniCode_Conv(STOCKREC.SAI_QTY, "000000000")                '差異数
                    Call UniCode_Conv(STOCKREC.FILLER, "")
                    If CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)) = 0 And _
                        CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)) = 0 Then
                    Else
                        Data_cnt = Data_cnt + 1
                    End If
                    com = BtOpGetNext
                Loop
            End If
        Next i
    Else
        
        
        '単独BU
        Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K6_ITEM.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K6_ITEM.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K6_ITEM.ST_DAN, Text(ptxS_DAN).Text)
        Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
        
        com = BtOpGetGreaterEqual
        
        Do
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                            
                    If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) > _
                        (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    
                    Exit Do
                
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "棚卸しデータ")
                    Exit Function
            End Select
                                            '棚卸しデータ作成
            Call UniCode_Conv(STOCKREC.JGYOBU, Last_JGYOBU)                             '事業部
            Call UniCode_Conv(STOCKREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))        '国内外
            Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '品目コード
                                                                                        '標準棚番
            Call UniCode_Conv(STOCKREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
            Call UniCode_Conv(STOCKREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                                                        '松下理論在庫
            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                
                    Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                    Exit Function
            End Select
            Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                                                                                    
            Call UniCode_Conv(STOCKREC.BU_ZAI_QTY, StrConv(SUMZREC.BU_ZAI_QTY, vbUnicode))
            Call UniCode_Conv(STOCKREC.PPSC_ZAI_QTY, StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode))
                                                                                    
                                                                                    
                                                                                    
                                                                                    
                                                                                    'ＰＯＳ総在庫集計
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                Exit Function
            End If
            Call UniCode_Conv(STOCKREC.POS_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
            POS_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                                    
            For i = 0 To UBound(Betu_Loc)
                Betu_Loc(i).SOKO = ""
                Betu_Loc(i).Retu = ""
                Betu_Loc(i).Ren = ""
                Betu_Loc(i).Dan = ""
                Betu_Loc(i).ZAIKO_QTY = 0
            Next i
            Betu_Zaiko_Qty = 0
                                                                                    '標準棚番在庫集計
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                    StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                
                Exit Function
            End If
            
            Call UniCode_Conv(STOCKREC.ST_ZAIKO, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
            
            Call UniCode_Conv(STOCKREC.EE1_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE1_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.EE2_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE2_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
            Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
            Call UniCode_Conv(STOCKREC.ETC_ZAIKO, "00000000")
            
            
            Betu_Loc(0).SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Betu_Loc(0).Retu = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Betu_Loc(0).Ren = StrConv(ITEMREC.ST_REN, vbUnicode)
            Betu_Loc(0).Dan = StrConv(ITEMREC.ST_DAN, vbUnicode)
            Betu_Loc(0).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
            
            Betu_Zaiko_Qty = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                                                                
            If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                    '別置き検索
            Else
                For i = 1 To UBound(Betu_Loc)
                
                    If Tana_Kensaku(Betu_Location) Then
                        Exit Function
                    End If
            
                    If Len(Trim(Betu_Location)) = 0 Then
                                                    '検索完了
                        Exit For
                    End If
                        
                    If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                    Betu_Location) Then
                        Exit Function
                    End If
                
                    Betu_Loc(i).SOKO = Left(Betu_Location, 2)
                    Betu_Loc(i).Retu = Mid(Betu_Location, 3, 2)
                    Betu_Loc(i).Ren = Mid(Betu_Location, 5, 2)
                    Betu_Loc(i).Dan = Right(Betu_Location, 2)
                    Betu_Loc(i).ZAIKO_QTY = Sumi_Zaiko_Qty + Mi_Zaiko_Qty
                    
                    Betu_Zaiko_Qty = Betu_Zaiko_Qty + (Sumi_Zaiko_Qty + Mi_Zaiko_Qty)
            
            
                    If POS_Zaiko_Qty = Betu_Zaiko_Qty Then
                                                '検索完了
                        Exit For
                    End If
            
                Next i
            
            
                                                '別置き１
                If Betu_Loc(1).ZAIKO_QTY <> 0 Then
                    Call UniCode_Conv(STOCKREC.EE1_LOCATION, (Betu_Loc(1).SOKO & Betu_Loc(1).Retu & Betu_Loc(1).Ren & Betu_Loc(1).Dan))
                    Call UniCode_Conv(STOCKREC.EE1_ZAIKO, Format(Betu_Loc(1).ZAIKO_QTY, "00000000"))
                End If
                                                '別置き２
                If Betu_Loc(2).ZAIKO_QTY <> 0 Then
                    Call UniCode_Conv(STOCKREC.EE2_LOCATION, (Betu_Loc(2).SOKO & Betu_Loc(2).Retu & Betu_Loc(2).Ren & Betu_Loc(2).Dan))
                    Call UniCode_Conv(STOCKREC.EE2_ZAIKO, Format(Betu_Loc(2).ZAIKO_QTY, "00000000"))
                End If
                                                '別置き３
                Call UniCode_Conv(STOCKREC.EE3_LOCATION, "")
                Call UniCode_Conv(STOCKREC.EE3_ZAIKO, "00000000")
                                                '別置き４
                Call UniCode_Conv(STOCKREC.ETC_ZAIKO, Format((POS_Zaiko_Qty - Betu_Zaiko_Qty), "00000000"))
            End If
            
            Call UniCode_Conv(STOCKREC.CHECK_MARK, "")                      '照合マーク
            Call UniCode_Conv(STOCKREC.PRINT_YMD, Format(Now, "YYYYMMDD"))  '印刷日付
            Call UniCode_Conv(STOCKREC.INPUT_YMD, "")                       '入力日付
            Call UniCode_Conv(STOCKREC.SAI_QTY, "000000000")                '差異数
            
           
            
            Call UniCode_Conv(STOCKREC.FILLER, "")
        
            If CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)) = 0 And _
                CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)) = 0 Then
            Else
            
                Data_cnt = Data_cnt + 1
            
            End If
        
            com = BtOpGetNext
        
        Loop

    End If

    Call Input_UnLock
    
    New_Count_Proc = Data_cnt

End Function


Private Function Count_Proc() As Long
'----------------------------------------------------------------------------
'                   棚卸し表印刷処理
'           2018.04.11
'----------------------------------------------------------------------------
Dim Lcnt        As Integer

Dim sts         As Integer
Dim com         As Integer

Dim Save_Soko   As String * 2

Dim Edit        As String

Dim X_Tab       As Integer

Dim Print_cnt   As Long


    Count_Proc = True

    Call Input_Lock


    Lcnt = LMAX

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    Print_cnt = 0           '2007.12.03
    
    
    
    
    If Last_JGYOBU = "*" Then
        '全BU
    
        Call UniCode_Conv(K4_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        
        Call UniCode_Conv(K4_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K4_STOCK.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K4_STOCK.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K4_STOCK.ST_DAN, Text(ptxS_DAN).Text)
        
        Call UniCode_Conv(K4_STOCK.HIN_GAI, "")
        
        com = BtOpGetGreaterEqual
        
        
        Do
            sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K4_STOCK, Len(K4_STOCK), 4)
            Select Case sts
                Case BtNoErr
                     If StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                                            
                    If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                        > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                    
                Case BtErrEOF
                    Exit Do
                    
                    
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                    Exit Function
            End Select
    '-------------------------------------------------  明細印刷
            If com = BtOpGetGreaterEqual Then
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
            End If
            
            If Save_Soko <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
                
            End If
            
            
            
            Print_cnt = Print_cnt + 1       '2007.12.03
            
            
            com = BtOpGetNext
        
        Loop
    
    
    
    
    
    Else
        '単独BU
        Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        
        Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K1_STOCK.ST_RETU, Text(ptxS_RETU).Text)
        Call UniCode_Conv(K1_STOCK.ST_REN, Text(ptxS_REN).Text)
        Call UniCode_Conv(K1_STOCK.ST_DAN, Text(ptxS_DAN).Text)
        
        Call UniCode_Conv(K1_STOCK.HIN_GAI, "")
        
        com = BtOpGetGreaterEqual
        
        
        Do
            sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
            Select Case sts
                Case BtNoErr
                    If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                        Exit Do
                    End If
                                            
                    If (StrConv(STOCKREC.ST_SOKO, vbUnicode) & StrConv(STOCKREC.ST_RETU, vbUnicode) & StrConv(STOCKREC.ST_REN, vbUnicode) & StrConv(STOCKREC.ST_DAN, vbUnicode)) _
                        > (Text(ptxSOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                        Exit Do
                    End If
                    
                Case BtErrEOF
                    Exit Do
                    
                    
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "棚卸しデータ")
                    Exit Function
            End Select
    '-------------------------------------------------  明細印刷
            If com = BtOpGetGreaterEqual Then
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
            End If
            
            If Save_Soko <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
                
            End If
            
            
            
            Print_cnt = Print_cnt + 1       '2007.12.03
            
            
            com = BtOpGetNext
        
        Loop


    End If


    Call Input_UnLock
    
    Count_Proc = Print_cnt

End Function


