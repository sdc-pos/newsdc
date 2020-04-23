VERSION 5.00
Begin VB.Form F1030551 
   BackColor       =   &H00FFFFFF&
   Caption         =   "大阪ＰＣ向け品番別出庫表印刷"
   ClientHeight    =   5655
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
   ScaleHeight     =   5655
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3825
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1020
      Width           =   285
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   4725
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2760
      Width           =   480
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   3885
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2760
      Width           =   480
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1620
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1620
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1620
      Width           =   615
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0=未出庫/9=出庫完了"
      Height          =   255
      Index           =   6
      Left            =   4275
      TabIndex        =   34
      Top             =   1140
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷件数/読込件数"
      Height          =   255
      Index           =   2
      Left            =   1530
      TabIndex        =   33
      Top             =   3540
      Width           =   2085
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出庫区分"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   32
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      Height          =   315
      Index           =   1
      Left            =   4680
      TabIndex        =   31
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   4
      Left            =   4545
      TabIndex        =   30
      Top             =   3540
      Width           =   120
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      Height          =   315
      Index           =   0
      Left            =   3915
      TabIndex        =   29
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   4410
      TabIndex        =   28
      Top             =   2880
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷頁範囲"
      Height          =   255
      Index           =   0
      Left            =   2415
      TabIndex        =   27
      Top             =   2880
      Width           =   1275
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
      TabIndex        =   26
      Top             =   5160
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "便"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   25
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   2
      Left            =   4920
      TabIndex        =   24
      Top             =   1740
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   1
      Left            =   4440
      TabIndex        =   23
      Top             =   1740
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷区分"
      Height          =   255
      Index           =   3
      Left            =   2730
      TabIndex        =   22
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷予定日"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030551"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxPRINT_KBN% = 0             '印刷区分


Private Const ptxSYUKA_YY% = 1              '出荷予定日　年
Private Const ptxSYUKA_MM% = 2              '出荷予定日　月
Private Const ptxSYUKA_DD% = 3              '出荷予定日　日

Private Const ptxINS_BIN% = 4               '便

Private Const ptxS_Page% = 5                '印刷開始　頁数
Private Const ptxE_Page% = 6                '印刷終了　頁数


Private Const Text_Max% = 6                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbPRINT_KBN% = 0            '印刷区分


Private Const Print_KBN0$ = "新　規"
Private Const Print_KBN1$ = "再印刷"
Private Const Print_KBN_SIN$ = "0"
Private Const Print_KBN_SAI$ = "1"

Private KASO_NYUKA_SOKO As String * 2       '仮想　入荷倉庫番号
Private KASO_SYOHN_SOKO As String * 2       '仮想　商品化倉庫番号
Private KASO_NAI_SOKO As String * 2         '仮想　内職倉庫番号


Private Const LMAX% = 40                    '頁内最大行数
Private Const MGN_L% = 2                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）
'Dim PRT_CAN As Boolean                     '印刷途中キャンセル要求


Private NormalFont As New StdFont           '印刷フォント
Private SmallFont As New StdFont            '印刷フォント   '2007.03.29

Private BoldFont As New StdFont             '印刷フォント   '2007.03.29

Private RargeFont As New StdFont             '印刷フォント   '2007.03.29



Private Code39Font As New StdFont           '印刷フォント



'ステーション№
Private WS_NO       As String * 10


Private Input_Cnt   As Integer
Private Output_Cnt  As Integer



Private P_CNT       As Integer
Private Print_F     As Boolean
Private Const LAST_UPDATE_DAY$ = "[F103055]2010.04.07 11:00"

Private Function Print_Proc() As Integer

Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 1

Dim com             As Integer
Dim sts             As Integer
    
Dim RetBuf          As String
    
Dim Print_F         As Boolean
    
    
    Print_Proc = True

    Call Input_Lock
    
    
    Lcnt = 99
    P_CNT = 0
    Print_F = False
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    Call UniCode_Conv(K1_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text)
    Call UniCode_Conv(K1_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_SOKO, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_RETU, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_REN, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_DAN, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.JGYOBU, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.NAIGAI, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.HIN_NO, "")
    
    
        
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K1_Y_SYU_SUM, Len(K1_Y_SYU_SUM), 1)
    
        Select Case sts
            Case BtNoErr
                
                If StrConv(Y_SYU_SUMREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                    Text(ptxSYUKA_MM).Text & _
                                                                    Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
                
                
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(大阪PC出庫表用)データ")
                Call Input_UnLock
                Exit Function
        End Select
                                            
        Input_Cnt = Input_Cnt + 1
        Label3(1).Caption = Format(Input_Cnt, "#0")
                                            
                                            
        Print_F = True
                                            
                                            
        Select Case Trim(Text(ptxPRINT_KBN).Text)
            Case ""
            Case "0"
                If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) = 0 Then
                    Print_F = False
                End If
            Case "9"
                If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) <> 0 Then
                    Print_F = False
                End If
        End Select
                                            
                                            
                                            
'        If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) = 0 Then
        
        If Not Print_F Then
        Else
            If Lcnt = 99 Then
                SAVE_SOKO_No = Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1)
            Else
                                                '倉庫のブレーク
                If SAVE_SOKO_No <> Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1) Then
                    Lcnt = LMAX + 1
                    SAVE_SOKO_No = Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1)
                End If
            End If

            If Lcnt > LMAX Then                 'ヘッダーコントロール
                If Head_Proc(Lcnt) Then
                    Call Input_Lock
                    Exit Function
                End If
            
                '指定ページ数OVER
                If P_CNT > CInt(Text(ptxE_Page).Text) Then
                    If Lcnt <> 99 Then
                        Printer.EndDoc
                    End If
                
                
                
                    Call Input_UnLock
                
                    Print_Proc = False
                
                    Exit Function
                
                End If
            
            
            End If
                                                
                                                            
            If P_CNT < CInt(Text(ptxS_Page).Text) Then
            Else
                                                
                                                
                Output_Cnt = Output_Cnt + 1
                Label3(0).Caption = Format(Output_Cnt, "#0")
                                                
                                                
                '-----------------------------------------------------  '１行目
                Printer.Print Tab(MGN_L);
                '標準棚番
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_RETU, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_REN, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_DAN, vbUnicode), 1, 2);
                '品番(外)
'                Printer.Print Tab(MGN_L + 14);
                Printer.Print Tab(MGN_L + 18);
                Set Printer.Font = RargeFont
                Printer.Print Left(StrConv(Y_SYU_SUMREC.HIN_NO, vbUnicode), 13);
                Set Printer.Font = NormalFont
    
    '            Printer.Print Tab(MGN_L + 34);
    '            '別置棚検索
    '            If Trim(StrConv(Y_SYU_SUMREC.BETU_SOKO, vbUnicode)) = "" Then
    '            Else
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_RETU, vbUnicode), 1, 2) & "-";
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_REN, vbUnicode), 1, 2) & "-";
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_DAN, vbUnicode), 1, 2);
    '            End If
    '            Printer.Print Tab(MGN_L + 46);
    '            '別置き在庫数
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.BETU_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
    '            Printer.Print Tab(MGN_L + 55);
    '            '商品化＆内職在庫数
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.SYO_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
    '            '入荷倉庫在庫
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.NYU_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
        
                '出庫済み／出庫予定数
'                Printer.Print Tab(MGN_L + 29);
                Printer.Print Tab(MGN_L + 40);
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Set Printer.Font = BoldFont
                Printer.Print RetBuf;
                Set Printer.Font = NormalFont
                
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print " (" & RetBuf & ")";
                
                
                
                
                '印刷フォント設定（Ｃｏｄｅ３９）
'                Printer.Print Tab(MGN_L + 51);
'                Set Printer.Font = Code39Font
'                                    'バーコード(*伝票ID*)
'                Printer.Print "*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*";
'                                                        '印刷フォント設定（通常）
'                Set Printer.Font = NormalFont
                
'                Printer.Print Tab(MGN_L + 75);
                Printer.Print Tab(MGN_L + 65);
'                Printer.Print "(*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*)";
                Printer.Print StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode);

                '標準棚　在庫数
                Printer.Print Tab(MGN_L + 93);
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.ST_ZAIKO_QTY, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Set Printer.Font = SmallFont
                Printer.Print RetBuf;
                Set Printer.Font = NormalFont
                
                '件数
                Printer.Print Tab(MGN_L + 105);
                RetBuf = Format(CInt(StrConv(Y_SYU_SUMREC.DATA_CNT, vbUnicode)), "#0")
                If Len(RetBuf) < 4 Then
                    RetBuf = Space(4 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf
                
                
                
                
                '標準棚番BC
                Printer.Print Tab(MGN_L);
                
                
                If Trim(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode)) <> "" Then
                    Set Printer.Font = Code39Font
                    Printer.Print "*/" & StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_RETU, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_REN, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_DAN, vbUnicode) & "*";
                    Set Printer.Font = NormalFont
                End If
                Printer.Print Tab(MGN_L + 65);
                                    'バーコード(*伝票ID*)
                Set Printer.Font = Code39Font
                Printer.Print "*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*"
                                                        '印刷フォント設定（通常）
                Set Printer.Font = NormalFont
                
                        
                Printer.Print String(115, "─")
'                Printer.Print
            
            End If
            
            
            
            
            
            Lcnt = Lcnt + 3
        
        End If
        
        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If

    Label3(0).Caption = Format(Output_Cnt, "#0")


    Call Input_UnLock

    Print_Proc = False

End Function
                                    
Private Function Head_Proc(Lcnt As Integer) As Integer
Dim i As Integer
Dim sts As Integer

    Head_Proc = True


    P_CNT = P_CNT + 1

    If P_CNT < CInt(Text(ptxS_Page).Text) Or P_CNT > CInt(Text(ptxE_Page).Text) Then
        Lcnt = 8 + MGN_U
        Head_Proc = False
        Exit Function
    End If


    If Print_F Then
        Printer.NewPage
    End If

    Print_F = True



    For i = 1 To MGN_U
        Printer.Print
    Next i

''    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    
    Printer.Print Tab(MGN_L + 31);
    
    '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
    If Trim(Text(ptxINS_BIN).Text) = "" Then
        Printer.Print "『便指定なし』出庫表";
    Else
        Printer.Print "『" + Text(ptxINS_BIN).Text + "便』出庫表";
    End If
    '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
    
    If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SAI Then
        Printer.Print Tab(MGN_L + 63);
        Printer.Print "【再印刷】";
    End If
    
    Printer.Print Tab(MGN_L + 81);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "  P." & Format(P_CNT, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L + 5);
    Printer.Print "倉庫：";
    Printer.Print Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1);
''    Printer.Print Tab(MGN_L + 15);
''    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode))
''    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''    Select Case sts
''        Case BtNoErr
''            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
''        Case BtErrKeyNotFound
''        Case Else
''            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
''            Exit Function
''    End Select
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
'    Printer.Print Tab(MGN_L + 14);
    Printer.Print Tab(MGN_L + 18);
    Printer.Print "品番（外部）";
'    Printer.Print Tab(MGN_L + 28);
    Printer.Print Tab(MGN_L + 39);
'    Printer.Print Tab(MGN_L + 35);
'    Printer.Print "別置棚番";
'    Printer.Print Tab(MGN_L + 47);
'    Printer.Print "別置在庫";
'    Printer.Print Tab(MGN_L + 56);
'    Printer.Print "商品化室";
'    Printer.Print Tab(MGN_L + 65);
'    Printer.Print "入荷倉庫";
'    Printer.Print Tab(MGN_L + 40);
    Set Printer.Font = BoldFont
    Printer.Print "出荷指示数";
    Set Printer.Font = NormalFont
    
'    Printer.Print Tab(MGN_L + 39);
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "(出庫 済数)";
    
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "出庫ＩＤ";
    
    
    Printer.Print Tab(MGN_L + 92);
    Set Printer.Font = SmallFont
    Printer.Print "標準棚在庫";
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 104);
    Printer.Print "件数"
    
    
    Printer.Print String(115, "─")
    
    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYU_HREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Tana_Kensaku = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    
    
    
    
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYU_HREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.SOKO_NO, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(Y_SYU_HREC.JGYOBU, vbUnicode) Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYU_HREC.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.SOKO_NO, vbUnicode) <> StrConv(ITEMREC.ST_SOKO, vbUnicode) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(ITEMREC.ST_RETU, vbUnicode) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(ITEMREC.ST_REN, vbUnicode) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(ITEMREC.ST_DAN, vbUnicode) Then
                                                'システム倉庫の判定
                    Call UniCode_Conv(K0_SOKO.SOKO_NO, StrConv(ZAIKOREC.SOKO_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '考えられないので読み飛ばし
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "倉庫マスタ")
                            Exit Function
                    End Select
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
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030551.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030551)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030551)


    F1030551.MousePointer = vbDefault

End Sub





Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbPRINT_KBN      '印刷区分
            Text(ptxPRINT_KBN).SetFocus
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim sts     As Integer
Dim com     As Integer
    
Dim c       As String * 128
    
    Select Case Index
        Case 8                              '印刷
            
            
            If Text(ptxPRINT_KBN).Text <> "0" And Text(ptxPRINT_KBN).Text <> "9" And Trim(Text(ptxPRINT_KBN).Text) <> "" Then
                MsgBox "入力内容はエラーです。"
                Text(ptxPRINT_KBN).SetFocus
                Exit Sub
            End If
            
'>>>>>>>>>>>>>>>>>>>>>  便№　ＡＬＬ指定可  2012.12.21
            
'2012.12.21            For i = ptxSYUKA_YY To ptxINS_BIN
            For i = ptxSYUKA_YY To ptxSYUKA_MM
            
'>>>>>>>>>>>>>>>>>>>>>  便№　ＡＬＬ指定可  2012.12.21
                
                
                
                
                If Not IsNumeric(Text(i).Text) Then
                    MsgBox "入力内容はエラーです。"
                    Text(i).SetFocus
                    Exit Sub
                End If
            
            
                Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
                            
            
            
            Next i
            
            
'>>>>>>>>>>>>>>>>>>>>>  便№　ＡＬＬ指定可  2012.12.21
            If Trim(Text(ptxINS_BIN).Text) = "" Then
            Else
                If Not IsNumeric(Text(ptxINS_BIN).Text) Then
                    MsgBox "入力内容はエラーです。(便№：数値 or 空白)"
                    Text(ptxINS_BIN).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>  便№　ＡＬＬ指定可  2012.12.21
            
            
            If Not IsNumeric(Text(ptxS_Page).Text) Then
                Text(ptxS_Page).Text = "001"
            Else
                Text(ptxS_Page).Text = Format(CInt(Text(ptxS_Page).Text), "000")
            End If
            
            If Not IsNumeric(Text(ptxE_Page).Text) Then
                Text(ptxE_Page).Text = "999"
            Else
                Text(ptxE_Page).Text = Format(CInt(Text(ptxE_Page).Text), "000")
            End If
            
            
            
            
            '印刷済みのﾁｪｯｸ
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                Call UniCode_Conv(K2_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                            Text(ptxSYUKA_MM).Text & _
                                                            Text(ptxSYUKA_DD).Text)
                Call UniCode_Conv(K2_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
                
                sts = BTRV(BtOpGetEqual, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                Select Case sts
                    Case BtNoErr
'                        MsgBox "指定内容は、出庫表印刷済みです。再印刷を選択して下さい。"
'                        Combo(pcmbPRINT_KBN).SetFocus
'                        Exit Sub
                                            
                        yn = MsgBox("指定内容は、出庫表印刷済みです。処理を継続しますか？", vbOKCancel + vbDefaultButton2 + vbCritical, "確認入力")
                        If yn = vbCancel Then
                            Combo(pcmbPRINT_KBN).SetFocus
                            Exit Sub
                        End If
                
                        Call UniCode_Conv(K2_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                                    Text(ptxSYUKA_MM).Text & _
                                                                    Text(ptxSYUKA_DD).Text)
                        Call UniCode_Conv(K2_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
                        com = BtOpGetGreaterEqual
                    
                        Do
                        
                            DoEvents
                        
                            sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                            Select Case sts
                                Case BtNoErr
                                
                                    If StrConv(Y_SYU_SUMREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                                    Text(ptxSYUKA_MM).Text & _
                                                                                    Text(ptxSYUKA_DD).Text) Then
                                        Exit Do
                                    End If
                                
                                    
                                    '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                                    If Trim(Text(ptxINS_BIN).Text) <> "" Then
                                        If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                                            Exit Do
                                        End If
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                                
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "出荷予定(大阪PC出庫表用)")
                                    Unload Me
                            End Select
                        
                        
                        
                            sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                            If sts Then
                                Call File_Error(sts, BtOpDelete, "出荷予定(大阪PC出庫表用)")
                                Unload Me
                            End If
                        
                            com = BtOpGetNext
                        
                        Loop
                    
                                            
                    
                    
                    
                    
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定(大阪PC出庫表用)")
                        Unload Me
                End Select
            End If
            
            
            
            
            
            
            
            Beep
            yn = MsgBox("「出庫表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                Input_Cnt = 0
                Output_Cnt = 0
                
                
                
                If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                    If Data_Make_Proc() Then
                        Unload Me
                    End If
                Else
                    If RE_Data_Make_Proc() Then
                        Unload Me
                    End If
                End If
                
                If Input_Cnt = 0 Then
                
                    MsgBox "「出庫表印刷」出庫対象が有りませんでした｡ "
                Else
                
                    Input_Cnt = 0
                    Output_Cnt = 0
                    
                    
                                    
                    
                    
                    If Print_Proc() Then
                        Unload Me
                    End If
                
                
                
                    If Output_Cnt = 0 Then
                        MsgBox "「出庫表印刷」出庫対象が有りませんでした｡ "
                    Else
                        MsgBox "「出庫表印刷」印刷終了しました。"
                
                    End If
                
                End If
            
            
            End If
            
            Combo(pcmbPRINT_KBN).SetFocus
                    
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
     
Dim sBuffer As String

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
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

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030551.Caption = "大阪ＰＣ向け品番別出庫表印刷（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
                                
                                '入荷仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '商品化仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '内職仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ファイルＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(大阪PC出庫表用)ファイルＯＰＥＮ
    If Y_SYU_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1030551.FontName
        .Size = 14
        .Bold = False
    End With
                                
                                '印刷フォント(小)設定   2007.03.29
    With SmallFont
        .NAME = F1030551.FontName
        .Size = 10
    End With
                                
                                '印刷フォント(小)設定   2007.03.29
    With BoldFont
        .NAME = F1030551.FontName
        .Size = 14
        .Bold = True
    End With
                                
                                
                                
    With RargeFont
        .NAME = F1030551.FontName
        .Size = 18
        .Bold = False
    End With
                                
                                
                                
                                '印刷フォント設定（バーコード）
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
                                
                                '画面初期設定
    
    
    Combo(pcmbPRINT_KBN).AddItem Print_KBN0 & "   " & Print_KBN_SIN
    Combo(pcmbPRINT_KBN).AddItem Print_KBN1 & "   " & Print_KBN_SAI
    Combo(pcmbPRINT_KBN).ListIndex = 0





    Text(ptxPRINT_KBN).Text = "0"


    Text(ptxSYUKA_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxSYUKA_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxSYUKA_DD).Text = Right(Format(Date, "yyyymmdd"), 2)


    Combo(pcmbPRINT_KBN).SetFocus
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
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
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030551 = Nothing

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
    F1030551.Caption = "大阪ＰＣ向け品番別出庫表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
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
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   大阪ＰＣ用出庫表データ作成処理(新規処理)
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim svJGYOBU    As String
Dim svNAIGAI    As String
Dim svHIN_NO    As String
    
Dim Y_SURYO     As Long
Dim J_SURYO     As Long
Dim DATA_CNT    As Integer
    
Dim SKIP_F      As Boolean
    
Dim ID_NO       As String * 12
    
    Data_Make_Proc = True
                                
    Call Input_Lock
                                
                                
                                            '出庫表データＣＬＯＳＥ
''    sts = BTRV(BtOpClose, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
''    If sts Then
''        If sts <> BtErrNoOpen Then
''            Call File_Error(sts, BtOpClose, "出荷予定(大阪PC出庫表用)データ")
''            Call Input_UnLock
''            Exit Function
''        End If
''    End If
''
''                                            '出庫表データＯＰＥＮ
''    If Y_SYU_SUM_Open(BtOpenNomal, WS_NO) Then
''        Call Input_UnLock
''        Exit Function
''    End If
                                            
                                            '前回値ｸﾘｱｰ
    
    Call UniCode_Conv(K3_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K3_Y_SYU_SUM.SYU_NO, "")
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
    
        Select Case sts
            Case BtNoErr
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(大阪PC出庫表用)データ")
                Call Input_UnLock
                Exit Function
        End Select






        sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "出荷予定(大阪PC出庫表用)データ")
            Call Input_UnLock
            Exit Function
        End If
    
        com = BtOpGetNext
    Loop
    
    '-----------------------------------    出荷予定ﾎｽﾄｲﾒｰｼﾞﾍﾞｰｽで集計処理開始
    svJGYOBU = ""
    svNAIGAI = ""
    svHIN_NO = ""
    
    Y_SURYO = 0
    J_SURYO = 0
    DATA_CNT = 0
    
    
    Input_Cnt = 0
    Output_Cnt = 0
    
    Call UniCode_Conv(K5_Y_SYU_H.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K5_Y_SYU_H.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))
    
    Call UniCode_Conv(K5_Y_SYU_H.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K5_Y_SYU_H.NAIGAI, NAIGAI_NAI)
    
    Call UniCode_Conv(K5_Y_SYU_H.HIN_NO, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K5_Y_SYU_H, Len(K5_Y_SYU_H), 5)
    
        Select Case sts
            Case BtNoErr
            
            
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
            
            
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                Text(ptxSYUKA_MM).Text & _
                                                                Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(大阪PC出庫表用)データ")
                Call Input_UnLock
                Exit Function
        End Select
            
        Input_Cnt = Input_Cnt + 1
    
        SKIP_F = False
        
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_F = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Call Input_UnLock
                Exit Function
        End Select
    
            
        If Not SKIP_F Then
    
            If Trim(svJGYOBU) = "" Then
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                sts = Den_No_Set_Proc(33, Last_JGYOBU, ID_NO)
                If sts Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
            
            End If
        
            If svJGYOBU <> StrConv(Y_SYU_HREC.JGYOBU, vbUnicode) Or _
                svNAIGAI <> StrConv(Y_SYU_HREC.NAIGAI, vbUnicode) Or _
                svHIN_NO <> StrConv(Y_SYU_HREC.HIN_NO, vbUnicode) Then
            
                If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, ID_NO) Then
                    Call Input_UnLock
                    Exit Function
                End If
            
                
                sts = Den_No_Set_Proc(33, Last_JGYOBU, ID_NO)
                If sts Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                Y_SURYO = 0
                J_SURYO = 0
                DATA_CNT = 0
            End If
    
    
        
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) <> "1" Then
            
                Y_SURYO = Y_SURYO + CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
                J_SURYO = J_SURYO + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
                DATA_CNT = DATA_CNT + 1
            
            End If
        
        
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts <> BtNoErr Then
        
                Call File_Error(sts, BtOpUpdate, "出荷予定データ")
                Call Input_UnLock
                Exit Function
            End If
        
            Call UniCode_Conv(Y_SYU_HREC.SYU_NO, ID_NO)
            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K5_Y_SYU_H, Len(K5_Y_SYU_H), 5)
            If sts <> BtNoErr Then
        
                Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                Call Input_UnLock
                Exit Function
            End If
        End If
    
        com = BtOpGetNext
    
    
    Loop
    
    If Trim(svJGYOBU) <> "" Then
        If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, ID_NO) Then
            Call Input_UnLock
            Exit Function
        End If
    End If
    
    
    Call Input_UnLock
    
    Data_Make_Proc = False
                                

End Function
Private Function RE_Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   大阪ＰＣ用出庫表データ作成処理(再印刷処理)
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim svID_NO     As String
Dim svJGYOBU    As String
Dim svNAIGAI    As String
Dim svHIN_NO    As String
    
    
Dim Y_SURYO     As Long
Dim J_SURYO     As Long
Dim DATA_CNT    As Integer
    
Dim SKIP_F      As Boolean
    
    RE_Data_Make_Proc = True
                                
    Call Input_Lock
                                
                                
                                            '出庫表データＣＬＯＳＥ
''    sts = BTRV(BtOpClose, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
''    If sts Then
''        If sts <> BtErrNoOpen Then
''            Call File_Error(sts, BtOpClose, "出荷予定(大阪PC出庫表用)データ")
''            Call Input_UnLock
''            Exit Function
''        End If
''    End If
''
''                                            '出庫表データＯＰＥＮ
''    If Y_SYU_SUM_Open(BtOpenNomal, WS_NO) Then
''        Call Input_UnLock
''        Exit Function
''    End If
                                            
                                            '前回値ｸﾘｱｰ
    Call UniCode_Conv(K3_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K3_Y_SYU_SUM.SYU_NO, "")
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
    
        Select Case sts
            Case BtNoErr
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(大阪PC出庫表用)データ")
                Call Input_UnLock
                Exit Function
        End Select


        sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "出荷予定(大阪PC出庫表用)データ")
            Call Input_UnLock
            Exit Function
        End If
    
        com = BtOpGetNext
    
    Loop
    
    '-----------------------------------    出荷予定ﾎｽﾄｲﾒｰｼﾞﾍﾞｰｽで集計処理開始
    
    svID_NO = ""
    
    Y_SURYO = 0
    J_SURYO = 0
    DATA_CNT = 0
    
    
    Call UniCode_Conv(K7_Y_SYU_H.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))
    Call UniCode_Conv(K7_Y_SYU_H.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K7_Y_SYU_H.SYU_NO, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K7_Y_SYU_H, Len(K7_Y_SYU_H), 7)
    
        Select Case sts
            Case BtNoErr
            
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                Text(ptxSYUKA_MM).Text & _
                                                                Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
            
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    便№　ＡＬＬ指定可  2012.12.21
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(大阪PC出庫表用)データ")
                Call Input_UnLock
                Exit Function
        End Select
    
        Input_Cnt = Input_Cnt + 1
    
    
        SKIP_F = False
        
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_F = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Call Input_UnLock
                Exit Function
        End Select
    
            
        If Not SKIP_F Then
    
            If Trim(svID_NO) = "" Then
            
                svID_NO = StrConv(Y_SYU_HREC.SYU_NO, vbUnicode)
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
            
            End If
        
            If svID_NO <> StrConv(Y_SYU_HREC.SYU_NO, vbUnicode) Then
            
                If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, svID_NO) Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
                svID_NO = StrConv(Y_SYU_HREC.SYU_NO, vbUnicode)
            
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                Y_SURYO = 0
                J_SURYO = 0
                DATA_CNT = 0
            End If
    
    
        
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) <> "1" Then
            
                Y_SURYO = Y_SURYO + CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
                J_SURYO = J_SURYO + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
                DATA_CNT = DATA_CNT + 1
            
            End If
    
        End If
    
        com = BtOpGetNext
    
    
    Loop
    
    If Trim(svID_NO) <> "" Then
        If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, svID_NO) Then
            Call Input_UnLock
            Exit Function
        End If
    End If
    
    
    Call Input_UnLock
    
    RE_Data_Make_Proc = False
                                

End Function

Private Function Data_Output_Proc(JGYOBU As String, NAIGAI As String, HIN_NO As String, _
                                    Y_SURYO As Long, J_SURYO As Long, DATA_CNT As Integer, ID_NO As String) As Integer
'----------------------------------------------------------------------------
'                   大阪ＰＣ用出庫表データ出力処理
'----------------------------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim TEMP_QTY        As Long

Dim Betu_LOCATION   As String * 8





    Data_Output_Proc = True





    Call UniCode_Conv(Y_SYU_SUMREC.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))

    Call UniCode_Conv(Y_SYU_SUMREC.INS_BIN, Text(ptxINS_BIN).Text)

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_NO)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            
            Data_Output_Proc = False
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select

    Call UniCode_Conv(Y_SYU_SUMREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))

    Call UniCode_Conv(Y_SYU_SUMREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_SYU_SUMREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_SYU_SUMREC.HIN_NO, HIN_NO)

    Call UniCode_Conv(Y_SYU_SUMREC.Y_SURYO, Format(Y_SURYO, "0000000"))
    Call UniCode_Conv(Y_SYU_SUMREC.J_SURYO, Format(J_SURYO, "0000000"))
        
    Call UniCode_Conv(Y_SYU_SUMREC.SYU_NO, ID_NO)

    Call UniCode_Conv(Y_SYU_SUMREC.DATA_CNT, Format(DATA_CNT, "0000"))

    '標準棚在庫数
    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        SUMI_QTY = 0
        MI_QTY = 0
    Else
        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode))) Then
            Exit Function
        End If
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.ST_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    '別置き棚番
    If Tana_Kensaku(Betu_LOCATION) Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_SOKO, Mid(Betu_LOCATION, 1, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_RETU, Mid(Betu_LOCATION, 3, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_REN, Mid(Betu_LOCATION, 5, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_DAN, Mid(Betu_LOCATION, 7, 2))
    If Trim(Betu_LOCATION) = "" Then
        SUMI_QTY = 0
        MI_QTY = 0
    Else
                                    '別置棚　在庫数
        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                Betu_LOCATION) Then
            Exit Function
        End If
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    '商品化室在庫数
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_SYOHN_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    TEMP_QTY = SUMI_QTY + MI_QTY
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_NAI_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    TEMP_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
    Call UniCode_Conv(Y_SYU_SUMREC.SYO_ZAIKO_QTY, Format(TEMP_QTY, "0000000"))
    '入荷倉庫在庫
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_NYUKA_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.NYU_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    
    Call UniCode_Conv(Y_SYU_SUMREC.INS_NOW, Format(Now, "YYYYMMDDHHMMSS"))
    
    Call UniCode_Conv(Y_SYU_SUMREC.FILLER, "")

    sts = BTRV(BtOpInsert, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrDuplicates
            
        Case Else
            Call File_Error(sts, BtOpInsert, "出荷予定(大阪PC出庫表用)データ")
            Exit Function
    End Select




    Data_Output_Proc = False

End Function






