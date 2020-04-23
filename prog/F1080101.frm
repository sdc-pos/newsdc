VERSION 5.00
Begin VB.Form F1080101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ＡＢＣ管理支援リスト印刷"
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
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   6120
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2520
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
      TabIndex        =   11
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
      TabIndex        =   10
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
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
      TabIndex        =   18
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   4440
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標準棚番　倉庫番号"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   15
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1080101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_Soko% = 0                '開始　倉庫
Private Const ptxE_Soko% = 1                '終了　倉庫

Private Const Text_Max% = 1                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNAIGAI% = 0               '国内外

Private Const LMAX% = 60                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate       As String                   '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime       As String                   '印刷開始時刻（ﾍｯﾀﾞｰ用）

Dim ABC_DATA    As String                   'ABCデータフルパス

Dim NormalFont  As New StdFont              '印刷フォント

Dim AVE_ZENKAI_YMD  As String              '月平均前回起動年月日時分秒

Dim CLEAR_MODE  As Boolean


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                  「ＡＢＣ管理支援リスト」印刷処理
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer
Dim Save_Soko       As String * 2

    Print_Proc = True

    Call Input_Lock         '画面項目ロック


    If Data_Make_Proc() Then
        Call Input_UnLock
        Exit Function
    End If


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORPortrait   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time

    Save_Soko = ""

    
    com = BtOpGetFirst
    LCNT = 99


    Do
        DoEvents

        sts = BTRV(com, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)

        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "ＡＢＣ管理集計データ")
                Exit Function
        End Select
        
        If Trim(StrConv(ABCREC.RANK_NOW, vbUnicode)) <> Trim(StrConv(ABCREC.RANK_NEW, vbUnicode)) Then
                                        
                                        
                                        'コントロールブレーク
            If LCNT = 99 Then
                Save_Soko = Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2)
            
            
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
            
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Exit Function
                End Select
            
            
            End If
                                        
            If Save_Soko <> Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2) Then
                LCNT = LMAX + 1
                Save_Soko = Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2)
                
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
            
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Exit Function
                End Select
                
            End If
            
                                        'ヘッダーコントロール
            If LCNT > LMAX Then
                Call Print_Head(LCNT)
            End If
            
            
            Printer.Print Tab(MGN_L);
            
            Printer.Print Mid(StrConv(ABCREC.ST_LOCATION, vbUnicode), 3, 2) & "-" & Mid(StrConv(ABCREC.ST_LOCATION, vbUnicode), 5, 2) & "-" & Mid(StrConv(ABCREC.ST_LOCATION, vbUnicode), 7, 2);
            
            Printer.Print Tab(MGN_L + 14);
            
            Printer.Print StrConv(ABCREC.PACKING_NO, vbUnicode);
            
            Printer.Print Tab(MGN_L + 28);
            
            Printer.Print StrConv(ABCREC.RANK_NOW, vbUnicode);
            
            
            Printer.Print Tab(MGN_L + 40);

            Printer.Print StrConv(ABCREC.HIN_GAI, vbUnicode);
            
            Printer.Print Tab(MGN_L + 78);

            Printer.Print StrConv(ABCREC.RANK_NEW, vbUnicode)

            LCNT = LCNT + 1
    
        End If
        
        com = BtOpGetNext
    
    Loop

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    Call Input_UnLock         '画面項目ロック解除

    Print_Proc = False
End Function

Private Sub Print_Head(LCNT As Integer)
'----------------------------------------------------------------------------
'                  ヘッダーコントロール処理
'----------------------------------------------------------------------------
                                        
Dim i       As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        'ヘッダー（１）
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    
    Printer.Print Tab(20);
    Printer.Print "＊＊＊  ＡＢＣ管理支援リスト  ＊＊＊";
    Printer.Print Tab(20);
    Printer.Print "（" & Trim(Left(Combo(pcmbNAIGAI).Text, Len(Combo(pcmbNAIGAI).Text) - 1)) & "）";
    Printer.Print Tab(60);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                                                          
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫№：";
    Printer.Print Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2) & " "; StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                                        
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 14);
    Printer.Print "箱 №";
    Printer.Print Tab(MGN_L + 25);
    Printer.Print "設定ランク";
    Printer.Print Tab(MGN_L + 40);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 74);
    Printer.Print "現在ランク"
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Function OUTPUT_Proc() As Integer
'----------------------------------------------------------------------------
'                  ＣＳＶデータ出力処理
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
Dim Ret             As Integer
    

Dim FileNo          As Integer
Dim fileName        As String

Dim Save_Soko       As String * 2

Dim Soko_No         As String * 2
Dim c               As String * 128


    OUTPUT_Proc = True
'実行中中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    If Data_Make_Proc() Then
        Call Input_UnLock
        Exit Function
    End If


    FileNo = FreeFile
    fileName = ABC_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo


    Write #FileNo, "ＡＢＣ管理支援リスト"
    Write #FileNo, "標準棚番", "箱№", "設定ランク", "品番", "現在ランク"

    
    com = BtOpGetFirst

    Save_Soko = ""


    Do
        DoEvents
        
        sts = BTRV(com, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)

        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "ＡＢＣ管理集計データ")
                Exit Function
        End Select
                                        
        If Trim(StrConv(ABCREC.RANK_NOW, vbUnicode)) <> Trim(StrConv(ABCREC.RANK_NEW, vbUnicode)) Then
                                                                                
                                        'コントロールブレーク
            If Len(Trim(Save_Soko)) = 0 Then
                Save_Soko = Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2)
        
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
            
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Exit Function
                End Select
        
                
                If GetIni("SOKO_NO", Save_Soko, "SYS", c) Then
                    Soko_No = Save_Soko
                Else
                    Soko_No = Trim(c)
                End If
                
                
                Write #FileNo, "倉庫:" & Soko_No & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode)
        
            End If
                                        
            If Save_Soko <> Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2) Then
                Save_Soko = Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2)
        
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
            
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Exit Function
                End Select
        
                If GetIni("SOKO_NO", Save_Soko, "SYS", c) Then
                    Soko_No = Save_Soko
                Else
                    Soko_No = Trim(c)
                End If
                
                
                Write #FileNo, "倉庫:" & Soko_No & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode)
        
            End If
                                                        '標準棚番
            Write #FileNo, Left(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2) & "-" & _
                            Mid(StrConv(ABCREC.ST_LOCATION, vbUnicode), 3, 2) & "-" & _
                            Mid(StrConv(ABCREC.ST_LOCATION, vbUnicode), 5, 2) & "-" & _
                            Right(StrConv(ABCREC.ST_LOCATION, vbUnicode), 2),
                                                        '箱№
            Write #FileNo, StrConv(ABCREC.PACKING_NO, vbUnicode),
                                                        '設定ランク
            Write #FileNo, StrConv(ABCREC.RANK_NOW, vbUnicode),
                                                        '品番
            Write #FileNo, StrConv(ABCREC.HIN_GAI, vbUnicode),
                                                        '新設定ランク
            Write #FileNo, StrConv(ABCREC.RANK_NEW, vbUnicode)
        
        End If

        com = BtOpGetNext
    Loop

    Close #FileNo
    
    
    
    
    
    
    
    
    Call Input_UnLock         '画面項目ロック解除
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
    End If
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1080101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1080101)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1080101)


    F1080101.MousePointer = vbDefault

End Sub

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   入力エラーチェック処理
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Err_Chk = True

    For i = ptxS_Soko To ptxE_Soko
        If IsNumeric(Text(i).Text) Then
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i
    
    If Trim(Text(ptxE_Soko).Text) = "" Then
        Text(ptxE_Soko).Text = "zz"
    End If
    
    If Text(ptxS_Soko).Text > Text(ptxE_Soko).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_Soko).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbNAIGAI        '注文区分
            Text(ptxS_Soko).SetFocus
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              'データ出力
            If Err_Chk() Then
                Exit Sub
            End If
        
            Beep
            ans = MsgBox("「ＡＢＣ管理支援リスト」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                
                
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「ＡＢＣ管理支援リスト」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
                    
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

Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer

Dim TUKI_AVE    As String

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
    LOG_F = Trim(c)
                                'ABCファイル名取り込み
    If GetIni("FILE", "ABC_DATA", "SYS", c) Then
        Beep
        MsgBox "ＡＢＣ管理ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ABC_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                '月平均出荷数プログラムＩＤ獲得
    If GetIni(App.EXEName, "TUKI_AVE", "SYS", c) Then
        Beep
        MsgBox "月平均出荷数ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    TUKI_AVE = Trim(c)
                                '月平均出荷数起動日付獲得
    If GetIni(TUKI_AVE, "ZENKAI_YMD", "SYS", c) Then
        AVE_ZENKAI_YMD = ""
    Else
        AVE_ZENKAI_YMD = Trim(c)
    End If





    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1080101.Caption = "ＡＢＣ管理支援リスト印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                               
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚別個装箱マスタＯＰＥＮ
    If TPACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '個装箱マスタＯＰＥＮ
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ＡＢＣ管理集計ファイルＯＰＥＮ
    If ABC_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1080101.FontName
        .Size = F1080101.FontSize
    End With
    Set Printer.Font = NormalFont
                                

    Show
                                
                                '画面初期設定
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbNAIGAI).SetFocus
    
    
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
                                            '棚別個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚別個装箱マスタ")
        End If
    End If
                                            '個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "個装箱マスタ")
        End If
    End If
                                            '月平均出荷数ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "個装箱マスタ")
        End If
    End If
                                            'ＡＢＣ管理集計ファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ＡＢＣ管理集計ファイル")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1080101 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).Code = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1080101.Caption = "ＡＢＣ管理支援リスト印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")"
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
                
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                  「ＡＢＣ管理集計ファイル」作成処理
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer
Dim ans                 As Integer

Dim RANK(0 To 5)        As Long
Dim i                   As Integer

Dim Upd_Com             As Integer

Dim c                   As String * 128

    Data_Make_Proc = True
'---------------------------------------------------------- 'レコードクリアー
    
    
                                '前回起動日獲得
    
                                '月平均出荷数起動日付獲得
    If GetIni(App.EXEName, "ZENKAI_YMD", "SYS", c) Then
        c = ""
    End If



    If AVE_ZENKAI_YMD > Trim(c) Then
        CLEAR_MODE = True       'クリアーする
    Else
        CLEAR_MODE = False      'クリアーしない
    End If
    
    
    
    If CLEAR_MODE Then
    
        com = BtOpGetFirst
        Do
            
            Do
                DoEvents
                sts = BTRV(com + BtSNoWait, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ABC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "ＡＢＣ管理集計ファイル")
                        Exit Function
                End Select
            Loop
        
            If sts = BtErrEOF Then
                Exit Do
            End If
            
            
            
            Do
                
                sts = BTRV(BtOpDelete, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ABC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "ＡＢＣ管理集計ファイル")
                        Exit Function
                End Select
            
            Loop
            
            com = BtOpGetNext
        
        Loop
    End If
'---------------------------------------------------------- '集計データ作成開始
    '品目マスタベースで処理開始
    Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxS_Soko).Text)
    Call UniCode_Conv(K6_ITEM.ST_RETU, "")
    Call UniCode_Conv(K6_ITEM.ST_REN, "")
    Call UniCode_Conv(K6_ITEM.ST_DAN, "")
    Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
            Select Case sts
                Case BtNoErr
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                        
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpUnlock, "品目マスタ")
                            Exit Function
                        End If
                        
                        sts = BtErrEOF
                    
                    End If
                    
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) > Trim(Text(ptxE_Soko).Text) Then
                        
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpUnlock, "品目マスタ")
                            Exit Function
                        End If
                        
                        
                        sts = BtErrEOF
                    End If
                
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                    Exit Function
            End Select
        Loop
                                                        
        If sts = BtErrEOF Then
            Exit Do
        End If
                                                        
                                                        
        If Len(Trim(StrConv(ITEMREC.PACKING_NO, vbUnicode))) = 0 Then
        
            sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpUnlock, "品目マスタ")
                Exit Function
            End If
        
        Else
                                                            
                                                            '事業部
            Call UniCode_Conv(K0_ABC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '国内外
            Call UniCode_Conv(K0_ABC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '標準棚番
            Call UniCode_Conv(K0_ABC.ST_LOCATION, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                            '箱№
            Call UniCode_Conv(K0_ABC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
                                                            '品番（外部）
            Call UniCode_Conv(K0_ABC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_Com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_Com = BtOpInsert
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ABC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "ＡＢＣ管理集計ファイル")
                        Exit Function
                End Select
            
            Loop
                                                            '事業部
            Call UniCode_Conv(ABCREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '国内外
            Call UniCode_Conv(ABCREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '標準棚番
            Call UniCode_Conv(ABCREC.ST_LOCATION, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                            '箱№
            Call UniCode_Conv(ABCREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
                                                            '品番（外部）
            Call UniCode_Conv(ABCREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    '                                                        '現在設定ランク(From 棚別個装箱マスタ)
    '        Call UniCode_Conv(K0_TPACKING.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
    '        Call UniCode_Conv(K0_TPACKING.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
    '        Call UniCode_Conv(K0_TPACKING.Ren, StrConv(ITEMREC.ST_RETU, vbUnicode))
    '        Call UniCode_Conv(K0_TPACKING.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
    '
    '        Do
    '
    '        sts = BTRV(BtOpGetEqual, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    '        Select Case sts
    '            Case BtNoErr
    '                Call UniCode_Conv(ABCREC.RANK_NOW, StrConv(TPACKINGREC.RANK, vbUnicode))
    '            Case BtErrKeyNotFound
    '                'ランク異常（棚／箱未登録）
    '                Call UniCode_Conv(ABCREC.RANK_NOW, "***")
    '            Case Else
    '                Call File_Error(sts, BtOpGetEqual, "棚別個装箱マスタ")
    '                Exit Function
    '        End Select
            If CLEAR_MODE Or _
                Upd_Com = BtOpInsert Then
                Call UniCode_Conv(ABCREC.RANK_NOW, StrConv(ITEMREC.RANK, vbUnicode))
            End If
                                                            '新ランク(From 月平均出荷数)
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "00000000")
                    Call UniCode_Conv(AVE_SYUKAREC.Two_Year_SYUKA, "00000000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
                    Exit Function
            End Select
                                                            '新ランク(From 個装箱マスタ)
            Call UniCode_Conv(K0_PACKING.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
            Select Case sts
                Case BtNoErr
                           
                    RANK(0) = CLng(StrConv(PACKINGREC.RANK_A1, vbUnicode))
                    RANK(1) = CLng(StrConv(PACKINGREC.RANK_A2, vbUnicode))
                    RANK(2) = CLng(StrConv(PACKINGREC.RANK_B1, vbUnicode))
                    RANK(3) = CLng(StrConv(PACKINGREC.RANK_B2, vbUnicode))
                    RANK(4) = CLng(StrConv(PACKINGREC.RANK_C1, vbUnicode))
                    RANK(5) = CLng(StrConv(PACKINGREC.RANK_C2, vbUnicode))
                
                            
                    For i = 0 To UBound(RANK)
                        
                        If CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)) > RANK(i) Then
                            Exit For
                        End If
                    
                    Next i
                
                    If i > UBound(RANK) Then
                        If CLng(StrConv(AVE_SYUKAREC.Two_Year_SYUKA, vbUnicode)) = 0 Then
                            Call UniCode_Conv(ABCREC.RANK_NEW, "E")
                        Else
                            Call UniCode_Conv(ABCREC.RANK_NEW, "D")
                        End If
                    Else
                        Select Case i
                            Case 0
                                Call UniCode_Conv(ABCREC.RANK_NEW, "A-1")
                            Case 1
                                Call UniCode_Conv(ABCREC.RANK_NEW, "A-2")
                            Case 2
                                Call UniCode_Conv(ABCREC.RANK_NEW, "B-1")
                            Case 3
                                Call UniCode_Conv(ABCREC.RANK_NEW, "B-2")
                            Case 4
                                Call UniCode_Conv(ABCREC.RANK_NEW, "C-1")
                            Case 5
                                Call UniCode_Conv(ABCREC.RANK_NEW, "C-2")
                        End Select
                    End If
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(ABCREC.RANK_NEW, "***")
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
                    Exit Function
            End Select
        
            Do
                sts = BTRV(Upd_Com, ABC_POS, ABCREC, Len(ABCREC), K0_ABC, Len(K0_ABC), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<ABC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "ＡＢＣ管理集計データ")
                        Exit Function
                End Select
            Loop
        
        
            If StrConv(ABCREC.RANK_NEW, vbUnicode) <> StrConv(ITEMREC.RANK, vbUnicode) Then
                                                    'ランク
        
                Call UniCode_Conv(ITEMREC.RANK, StrConv(ABCREC.RANK_NEW, vbUnicode))
        
                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "品目マスタ")
                            Exit Function
                    End Select
                Loop
        
        
        
            Else
        
                sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpUnlock, "品目マスタ")
                    Exit Function
                End If
            End If
        End If
        
        com = BtOpGetNext
    
    
    Loop
    
    If WriteIni(App.EXEName, "ZENKAI_YMD", "SYS", Format(Now, "YYYY/MM/DD HH:MM:SS")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & " ZENKAI_YMD")
        Exit Function
    End If
    
    
    Data_Make_Proc = False

End Function
