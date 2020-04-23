VERSION 5.00
Begin VB.Form F1090301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "在庫差異チェックリスト印刷"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "処理区分"
      Height          =   1215
      Left            =   4200
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "前日移動分"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "全　　　件"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷中断"
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
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   5160
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1200
      Width           =   855
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      TabIndex        =   16
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷中です"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   4200
      TabIndex        =   13
      Top             =   1320
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
Attribute VB_Name = "F1090301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0               '国内外

Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）
Dim SUMZAI_DATA  As String                  '在庫差異データフルパス

Dim NormalFont As New StdFont               '印刷フォント

Dim PRT_CAN As Boolean                      '印刷途中キャンセル要求
Private Function OUTPUT_Proc(Mode As Integer) As Integer

Dim com         As Integer
Dim sts         As Integer
Dim Ret         As Integer

Dim FileNo      As Integer
Dim fileName    As String
    
Dim Skip_Flg    As Boolean
    
    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock             '画面項目ロック

    FileNo = FreeFile
    fileName = SUMZAI_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo
    
    Write #FileNo, "品番（外）", "品名", "品番（内）", "ＰＣ在庫", , "ホスト在庫", , "差異数"
    
    Call UniCode_Conv(K1_SUMZ.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_SUMZ.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K1_SUMZ.ST_SOKO, "")
    Call UniCode_Conv(K1_SUMZ.ST_RETU, "")
    Call UniCode_Conv(K1_SUMZ.ST_REN, "")
    Call UniCode_Conv(K1_SUMZ.ST_DAN, "")
    Call UniCode_Conv(K1_SUMZ.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K1_SUMZ, Len(K1_SUMZ), 1)

        Select Case sts
            Case BtNoErr
                If StrConv(SUMZREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(SUMZREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                                            '範囲オーバー
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫集計データ")
                Exit Function
        End Select
        Skip_Flg = False
        
        If Mode = 1 And _
            CLng(StrConv(SUMZREC.SAI_QTY, vbUnicode)) = 0 Then
                                            '差異なしは印刷対象外
            Skip_Flg = True
        End If
        
        If Option1(1).Value Then
            If CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) = CLng(StrConv(SUMZREC.ZEN_Zai_Qty, vbUnicode)) Then
                                            '前日から差異なしは対象外
                Skip_Flg = True
            End If
        End If
        
        
        
        If Not Skip_Flg Then
                                        '明細印刷
            
            Write #FileNo, StrConv(SUMZREC.HIN_GAI, vbUnicode),
                                        '品目マスタ読込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
            Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
            Write #FileNo, Format(CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)), "#0"),
            If CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) < CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) Then
                Write #FileNo, "＜",
            Else
                Write #FileNo, ,
            End If
            Write #FileNo, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#0"),
            Write #FileNo, "＝",
            Write #FileNo, Format(CLng(StrConv(SUMZREC.SAI_QTY, vbUnicode)), "#0"),
            Write #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) + "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
    
        End If
    
        com = BtOpGetNext
    
    Loop

    Close #FileNo
    Call Input_UnLock             '画面項目ロック解除
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    OUTPUT_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function

Private Function Print_Proc() As Integer

Dim Lcnt        As Integer
Dim com         As Integer
Dim sts         As Integer
Dim RetBuf      As String
    
Dim Skip_Flg    As Boolean
    
    Print_Proc = True
'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock             '画面項目ロック
'    Label1.Visible = True
'    Command1.Visible = True
'    Command1.Enabled = True

'    PRT_CAN = False

    
    Lcnt = 99
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time
    
    Call UniCode_Conv(K1_SUMZ.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_SUMZ.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K1_SUMZ.ST_SOKO, "")
    Call UniCode_Conv(K1_SUMZ.ST_RETU, "")
    Call UniCode_Conv(K1_SUMZ.ST_REN, "")
    Call UniCode_Conv(K1_SUMZ.ST_DAN, "")
    Call UniCode_Conv(K1_SUMZ.HIN_GAI, "")
    
    
    com = BtOpGetGreater
    Do
        DoEvents
                                            '印刷中断要求
'        If PRT_CAN Then
'            Printer.KillDoc
'            Call Input_UnLock             '画面項目ロック
'            Label1.Visible = False
'            Command1.Visible = False
'            Print_Proc = False
'            Exit Function
'        End If
        
        sts = BTRV(com, SUMZ_POS, SUMZREC, Len(SUMZREC), K1_SUMZ, Len(K1_SUMZ), 1)

        Select Case sts
            Case BtNoErr
                If StrConv(SUMZREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(SUMZREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                                            '範囲オーバー
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫集計データ")
                Exit Function
        End Select
        
        Skip_Flg = False
        If CLng(StrConv(SUMZREC.SAI_QTY, vbUnicode)) = 0 Then
                                            '差異なしは印刷対象外
            Skip_Flg = True
        End If
                                           
        If Option1(1).Value Then
            If CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) = CLng(StrConv(SUMZREC.ZEN_Zai_Qty, vbUnicode)) Then
                Skip_Flg = True
            End If
        End If
                                           
        If Not Skip_Flg Then
                                           'ヘッダーコントロール
            If Lcnt > LMAX Then
                Call Print_Head(Lcnt)
            End If
                                        
                                        '明細印刷
            Printer.Print Tab(MGN_L);
            Printer.Print Left(StrConv(SUMZREC.HIN_GAI, vbUnicode), 13);
            Printer.Print Tab(MGN_L + 15);
                                        '品目マスタ読込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
            Printer.Print Tab(MGN_L + 45);
            Printer.Print Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 13);
            
            
            Printer.Print Tab(MGN_L + 60);
            RetBuf = Trim(Format(CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)), "#,##0"))
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            If CLng(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)) <> 0 Then
                Printer.Print "(";
                RetBuf = Trim(Format(CLng(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)), "#,##0"))
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                Printer.Print ")";
            End If



'If Trim(StrConv(SUMZREC.HIN_GAI, vbUnicode)) = "AMC55R-TD0V" Or _
'    Trim(StrConv(SUMZREC.HIN_GAI, vbUnicode)) = "AMC57B-UC0W" Then
'    Debug.Print StrConv(SUMZREC.HIN_GAI, vbUnicode) & "=" & RetBuf
'End If
            
            Printer.Print Tab(MGN_L + 80);
            If CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) < CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) Then
                Printer.Print "＜";
            End If
                        
            Printer.Print Tab(MGN_L + 86);
            RetBuf = Trim(Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#,##0"))
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            
            Printer.Print Tab(MGN_L + 101);
            Printer.Print "＝";
            
            Printer.Print Tab(MGN_L + 109);
            RetBuf = Trim(Format(CLng(StrConv(SUMZREC.SAI_QTY, vbUnicode)), "#,##0"))
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            
            If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                Printer.Print
            Else
                Printer.Print Space(1);
                Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-";
                Printer.Print StrConv(ITEMREC.ST_RETU, vbUnicode) & "-";
                Printer.Print StrConv(ITEMREC.ST_REN, vbUnicode) & "-";
                Printer.Print StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
            
            Lcnt = Lcnt + 1
    
        End If
    
        com = BtOpGetNext
    
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If
    
    Call Input_UnLock             '画面項目ロック解除
    Label1.Visible = False
    Command1.Visible = False
    
    Print_Proc = False

End Function

Private Sub Print_Head(Lcnt As Integer)
                                        
Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(26);
    Printer.Print "＊＊＊  在庫差異チェックリスト  ＊＊＊（";
    Printer.Print Trim(Left(Combo(pcmbNAIGAI).Text, Len(Combo(pcmbNAIGAI).Text) - 1));
    Printer.Print "）";
    Printer.Print Tab(101);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "品  名  ";
    Printer.Print Tab(MGN_L + 45);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 60);
    Printer.Print "ＰＣ在庫";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "ホスト在庫";
    Printer.Print Tab(MGN_L + 112);
    Printer.Print "差異数"
    Printer.Print

    Lcnt = 6 + MGN_U

End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1090301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1090301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1090301)


    F1090301.MousePointer = vbDefault

End Sub
Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              'データ
            
            Beep
            ans = MsgBox("「在庫差異チェックリスト」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                Beep
                ans = MsgBox("差異なしの品番も出力しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
                If ans = vbYes Then
                    If OUTPUT_Proc(0) Then
                        Unload Me
                    End If
                Else
                    If OUTPUT_Proc(1) Then
                        Unload Me
                    End If
                End If
'                Call Clear_Field
            End If
            Combo(pcmbNAIGAI).SetFocus
        Case 8                              '印刷
            
            Beep
            ans = MsgBox("「在庫差異チェックリスト」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Combo(pcmbNAIGAI).SetFocus
                    
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
   
End Sub
Private Sub Command1_Click()
'    PRT_CAN = True
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

Dim i   As Integer
Dim c   As String * 128
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
    LOG_F = Trim(c)
                                '在庫差異ファイル名取り込み
    If GetIni("FILE", "SUMZAI_DATA", "SYS", c) Then
        Beep
        MsgBox "在庫差異ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    SUMZAI_DATA = Trim(c)
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
            F1090301.Caption = "在庫差異チェックリスト印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1090301.FontName
        .Size = F1090301.FontSize
    End With
    Set Printer.Font = NormalFont
                                '画面初期設定
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Option1(0).Value = True
    Option1(1).Value = False
    
    Combo(pcmbNAIGAI).SetFocus
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
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
    Set F1090301 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1090301.Caption = "在庫差異チェックリスト印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
End Sub



