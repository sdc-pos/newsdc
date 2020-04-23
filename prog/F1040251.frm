VERSION 5.00
Begin VB.Form F1040251 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品番別在庫一覧表印刷（在庫なし含む）([F104025]2012.10.11 09:00)"
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
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   14
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1800
      Width           =   2535
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
      TabIndex        =   8
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
      TabIndex        =   20
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷中です"
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   17
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1040251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_HIN_GAI% = 0             '開始　品番
Private Const ptxE_HIN_GAI% = 1             '終了　品番

Private Const Text_Max% = 1                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNAIGAI% = 0               '国内外

Private Const LMAX% = 42                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate       As String                   '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime       As String                   '印刷開始時刻（ﾍｯﾀﾞｰ用）
Dim ZAIKO_DATA  As String                   '在庫データフルパス

Dim NormalFont  As New StdFont              '印刷フォント

Dim PRT_CAN     As Boolean                  '印刷途中キャンセル要求
Private Const Last_Update_Day$ = "[F104025]2012.10.11 09:00"


Private Function Print_Proc(Mode As Integer) As Integer
    
Dim sts             As Integer
Dim ZAIKO_com       As Integer
Dim ITEM_com        As Integer
    
Dim LCNT            As Integer
Dim Sum_Yuko_Z_Qty  As Long
Dim PRI_HIN_GAI     As String * 20
Dim RetBuf          As String

    Print_Proc = True
'印刷中は「印刷中断」以外のイベント取得不可
    Call Input_Lock         '画面項目ロック
    Label1.Visible = True
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    Pdate = Date
    Ptime = Time



    PRT_CAN = False
    LCNT = 99
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxS_HIN_GAI).Text)

    ITEM_com = BtOpGetGreaterEqual

    Do
        DoEvents
                                            '印刷中断要求
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '画面項目ロック解除
            Label1.Visible = False
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If
        
        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Or _
                    RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) > Text(ptxE_HIN_GAI).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "品目マスタ")
                Exit Function
        End Select

        Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K6_ZAIKO.Retu, "")
        Call UniCode_Conv(K6_ZAIKO.Ren, "")
        Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
        ZAIKO_com = BtOpGetGreater
        PRI_HIN_GAI = ""
        Sum_Yuko_Z_Qty = 0

        Do

            sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)

            Select Case sts
                Case BtNoErr
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, ZAIKO_com, "在庫データ")
                    Exit Function
            End Select
                                        'ヘッダーコントロール
            If LCNT > LMAX Then
                Call Print_Head(LCNT)
                PRI_HIN_GAI = ""
            End If
                                        '明細印刷
            If PRI_HIN_GAI <> StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
                Printer.Print Tab(MGN_L);
                Printer.Print StrConv(ZAIKOREC.HIN_GAI, vbUnicode);
                Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
            
                PRI_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
        
            End If
        
            Printer.Print Tab(MGN_L + 49);
            Printer.Print Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/";
            Printer.Print Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/";
            Printer.Print Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2);
        
            Printer.Print Tab(MGN_L + 66);
            Printer.Print StrConv(ZAIKOREC.HIN_NAI, vbUnicode);
        
            Printer.Print Tab(MGN_L + 81);
            Printer.Print StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-";
            Printer.Print StrConv(ZAIKOREC.Retu, vbUnicode) & "-";
            Printer.Print StrConv(ZAIKOREC.Ren, vbUnicode) & "-";
            Printer.Print StrConv(ZAIKOREC.Dan, vbUnicode);
        
            Printer.Print Tab(MGN_L + 97);
            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                Printer.Print "(済)";
            Else
                Printer.Print "(未)";
            End If
        
            RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            
            Printer.Print RetBuf;

            Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))

            Printer.Print Tab(MGN_L + 117);
            RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;

            Printer.Print
            LCNT = LCNT + 1

            ZAIKO_com = BtOpGetNext
        Loop
            
        If Sum_Yuko_Z_Qty = 0 Then
            If Mode = 0 Then
            
                    
                If LCNT > LMAX Then
                    Call Print_Head(LCNT)
                    PRI_HIN_GAI = ""
                End If
                    
                Printer.Print Tab(MGN_L);
                Printer.Print StrConv(ITEMREC.HIN_GAI, vbUnicode);
                Printer.Print StrConv(ITEMREC.HIN_NAME, vbUnicode);
            
                Printer.Print Tab(MGN_L + 117);
                RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;

                Printer.Print
                LCNT = LCNT + 1
            End If
        End If
            
        If LCNT > LMAX Then
            Call Print_Head(LCNT)
            PRI_HIN_GAI = ""
        End If
        
        Printer.Print
        LCNT = LCNT + 1
        
        ITEM_com = BtOpGetNext
    
    Loop

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    Call Input_UnLock         '画面項目ロック解除
    Label1.Visible = False
    Command1.Visible = False

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
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(36);
    Printer.Print "＊＊＊  品番別在庫一覧表  ＊＊＊ ";
    Printer.Print Tab(71);
    Printer.Print "（" & Trim(Left(Combo(pcmbNAIGAI).Text, Len(Combo(pcmbNAIGAI).Text) - 1)) & "）";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        
                                        '明細印刷
    Printer.Print Tab(MGN_L);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 21);
    Printer.Print "品  名  ";
    Printer.Print Tab(MGN_L + 49);
    Printer.Print "入荷日";
    Printer.Print Tab(MGN_L + 66);
    Printer.Print "品番（内部）";
    Printer.Print Tab(MGN_L + 81);
    Printer.Print "棚　番";
    Printer.Print Tab(MGN_L + 100);
    Printer.Print "有効在庫数";         '97.07.16
    Printer.Print Tab(MGN_L + 116);
    Printer.Print "累計在庫数"
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Function OUTPUT_Proc(Mode As Integer) As Integer
    
Dim sts             As Integer
Dim ZAIKO_com       As Integer
Dim ITEM_com        As Integer
Dim Ret             As Integer
    
Dim Sum_Yuko_Z_Qty  As Long
Dim PRI_HIN_GAI     As String * 13

Dim FileNo          As Integer
Dim fileName        As String

Dim c               As String * 128
Dim Soko_No         As String * 2

Dim BU_CNT          As Long


    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    FileNo = FreeFile
    fileName = ZAIKO_DATA
    
    
'2012.10.10    Ret = InStr(1, Trim(fileName), ".") - 1
    Ret = InStrRev(Trim(fileName), ".") - 1     '2012.10.10
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo


  

    Write #FileNo, "品番（外）", "品名", "入荷日", "品番（内）", "棚番", " ", "在庫数", "累計数"


    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxS_HIN_GAI).Text)

    ITEM_com = BtOpGetGreaterEqual

    BU_CNT = 0

    Do
        DoEvents
        
        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Or _
                    RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) > Text(ptxE_HIN_GAI).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "品目マスタ")
                Exit Function
        End Select

        Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K6_ZAIKO.Retu, "")
        Call UniCode_Conv(K6_ZAIKO.Ren, "")
        Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
        ZAIKO_com = BtOpGetGreater
        PRI_HIN_GAI = ""
        Sum_Yuko_Z_Qty = 0

        Do

            sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)

            Select Case sts
                Case BtNoErr
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, ZAIKO_com, "在庫データ")
                    Exit Function
            End Select
                                        '明細印刷
            If PRI_HIN_GAI <> StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
                Write #FileNo, StrConv(ZAIKOREC.HIN_GAI, vbUnicode),
                Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
            
                PRI_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            Else
                Write #FileNo, , ,
            End If
        
            Write #FileNo, Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2),
            Write #FileNo, StrConv(ZAIKOREC.HIN_NAI, vbUnicode),
        
            
            If GetIni("SOKO_NO", StrConv(ZAIKOREC.Soko_No, vbUnicode), "SYS", c) Then
                Soko_No = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            Else
                Soko_No = Trim(c)
            End If
            
            Write #FileNo, Soko_No & "-" & StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & StrConv(ZAIKOREC.Dan, vbUnicode),
            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                Write #FileNo, "（済）",
            Else
                Write #FileNo, "（未）",
            End If
            Write #FileNo, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0"),
            Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Write #FileNo, Format(Sum_Yuko_Z_Qty, "#0"),
            Write #FileNo,
''''''''''            Write #FileNo, "/" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
            

            ZAIKO_com = BtOpGetNext
        Loop
            
        If Sum_Yuko_Z_Qty = 0 Then
            If Mode = 0 Then
                    
                Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
                Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode), , , , ,
                Write #FileNo, Format(Sum_Yuko_Z_Qty, "#0")
            
            End If
        
        
        
        Else
        
        End If
            
        
        
        ITEM_com = BtOpGetNext
    
    Loop


    Close #FileNo
    
    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"


'''''''''''MsgBox "BU_CNT=" & Format(BU_CNT, "#")

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
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1040251.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040251)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040251)


    F1040251.MousePointer = vbDefault

End Sub

                                            'エラーチェック
Private Function Err_Chk() As Integer
                                            
                                            
                                            
    Err_Chk = True

'品番(外部)
    If Len(Text(ptxE_HIN_GAI).Text) = 0 Then
        Text(ptxE_HIN_GAI).Text = String(Len(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)), "z")
    End If

    If Text(ptxS_HIN_GAI).Text > Text(ptxE_HIN_GAI).Text Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxS_HIN_GAI).SetFocus
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
            Text(ptxS_HIN_GAI).SetFocus
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
            ans = MsgBox("「品番別在庫一覧表」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
                
            If ans = vbYes Then
                Beep
                ans = MsgBox("在庫なしの品番も出力しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
                    
                If ans = vbYes Then
                    If OUTPUT_Proc(0) Then
                        Unload Me
                    End If
                Else
                    If OUTPUT_Proc(1) Then
                        Unload Me
                    End If
                End If
            End If
            
            Combo(pcmbNAIGAI).SetFocus
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「品番別在庫一覧表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                Beep
                ans = MsgBox("在庫なしの品番も印刷しますか？", vbYesNo + vbQuestion, "確認入力")
                    
                If ans = vbYes Then
                    If Print_Proc(0) Then
                        Unload Me
                    End If
                Else
                    If Print_Proc(1) Then
                        Unload Me
                    End If
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
    LOG_F = Trim(c)
                                '在庫ファイル名取り込み
    If GetIni("FILE", "ZAIKO_DATA", "SYS", c) Then
        Beep
        MsgBox "在庫ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ZAIKO_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1040251.Caption = "品番別在庫一覧表印刷（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)       '2012.10.10

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1040251.FontName
        .Size = F1040251.FontSize
    End With
    Set Printer.Font = NormalFont
                                '画面初期設定
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbNAIGAI).SetFocus
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
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
    
    Set F1040251 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1040251.Caption = "品番別在庫一覧表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")" & Last_Update_Day$
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

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


