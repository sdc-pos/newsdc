VERSION 5.00
Begin VB.Form F1200101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ホスト棚番設定データ作成(V2.01)"
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
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   18
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1800
      Width           =   615
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
   Begin VB.Label LblFileName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   7140
      TabIndex        =   30
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label LblFileName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   7140
      TabIndex        =   29
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label LblFileName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7140
      TabIndex        =   28
      Top             =   2520
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出力件数"
      Height          =   375
      Index           =   5
      Left            =   2310
      TabIndex        =   27
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3675
      TabIndex        =   26
      Top             =   3480
      Width           =   1845
   End
   Begin VB.Label LblName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3675
      TabIndex        =   25
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Label LblName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3675
      TabIndex        =   24
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Label LblCnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5775
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label LblCnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5775
      TabIndex        =   22
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label LblCnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5775
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(指定した日付以降の品目を出力します。)"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   20
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   19
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   1920
      Width           =   255
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
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番設定日"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "F1200101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxYY% = 0                    '設定日　年
Private Const ptxMM% = 1                    '設定日　月
Private Const ptxDD% = 2                    '設定日　日

Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim Max_Soko    As Integer

Dim HTANA_DATA  As String                   'ホスト棚番設定データフルパス
Dim JGYOBA_CODE As String                   '事業場ｺｰﾄﾞ

'2005/05/31 Add Start #######################################################################
Dim OUT_SYUSI       As Variant              '(ini) 出力用収支 配列
'2005/05/31 Add End   #######################################################################

Private Function OUTPUT_Proc() As Integer
    
Dim sts                     As Integer
Dim ZAIKO_com               As Integer
Dim ITEM_com                As Integer


Dim Location1               As String
Dim Location2               As String
Dim Location3               As String


Dim Ret                     As Integer
Dim FileNo                  As Integer
Dim fileName                As String

Dim c                       As String * 128
Dim Soko_No                 As String * 2

Dim Count                   As Integer

Dim i                       As Integer
Dim j                       As Integer
Dim k                       As Integer




    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック


    For k = 0 To UBound(JGYOBU_T)

        Last_JGYOBU = JGYOBU_T(k).Code
        LblName(k).Caption = Trim(JGYOBU_T(k).NAME)



                                    '事業場ｺｰﾄﾞ取り込み
        If GetIni(App.EXEName, "JGYOBA_CODE" & "_" & Last_JGYOBU, "SYS", c) Then
        Else
            JGYOBA_CODE = Trim(c)
        
        
        
            FileNo = FreeFile
            fileName = HTANA_DATA
            
            Ret = InStr(1, Trim(fileName), ".") - 1
            fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
            
            On Error GoTo Error_Proc
            Open (fileName) For Output As FileNo
        '    Call UniCode_Conv(K3_ITEM.JGYOBU, Last_JGYOBU)
        '    Call UniCode_Conv(K3_ITEM.ST_SET_DT, Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text)
        
        
        
        
            Count = 0
        
            For i = 1 To 2
                
                                        '出力用収支 取り込み    2005/05/16
                If GetIni(App.EXEName, "SYUSI" & Last_JGYOBU & Format(i, "0"), "SYS", c) Then
                    Beep
                    MsgBox "出力用収支の獲得に失敗しました。処理を中止して下さい。"
                    End
                End If
                OUT_SYUSI = Split(Trim(c), ",", -1)
                
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, CStr(i))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
                ITEM_com = BtOpGetGreaterEqual
                Do
                    DoEvents
                    sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
                    Select Case sts
                        Case BtNoErr
        
                            If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                               StrConv(ITEMREC.NAIGAI, vbUnicode) <> CStr(i) Then
                                Exit Do
                            End If
        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call File_Error(sts, ITEM_com, "品目マスタ")
                            Exit Function
                    End Select
        
        
                    If StrConv(ITEMREC.ST_SET_DT, vbUnicode) < (Text(0).Text & Text(1).Text & Text(2).Text) Then
                    Else
                        
                        
                        
                        Select Case Last_JGYOBU
                            Case SENTAKU
                                Location1 = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                Location2 = ""
                                Location3 = ""


                                Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
                                Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                        


                                ZAIKO_com = BtOpGetGreater

                                Do
                                    DoEvents
                        
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
                        
                        
                                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) = Location1 Then
                                                
                                    Else
                                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                                                Else
                                                    If Location2 = "" Then
                                                        Location2 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                                    Else
                                                        
                                                        If Location2 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                                        Else
                                                            Location3 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                                            Exit Do
                                                        End If
                                                    End If
                                                End If
                                            
                                            
                                            
                                            Case BtErrKeyNotFound
                                                
                                                MsgBox "[倉庫ﾏｽﾀ異常]  Soko_No = " & StrConv(ZAIKOREC.Soko_No, vbUnicode)
                                                
                                                Exit Do
                                            Case Else
                                                Call File_Error(sts, ZAIKO_com, "在庫データ")
                                                Exit Function
                                        End Select
                                    End If
                                
                                    ZAIKO_com = BtOpGetNext
                                
                                Loop
                            
                            
                            Case Else
                                Location1 = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                    
                                Location2 = StrConv(ITEMREC.GLICS2_TANA, vbUnicode)
                    
                                Location3 = StrConv(ITEMREC.GLICS3_TANA, vbUnicode)
                    
                        End Select
                    
                        If Trim(Location1) = "" And Trim(Location2) = "" And Trim(Location3) = "" Then
                        Else
                    
            
                            For j = 0 To UBound(OUT_SYUSI)
                                Print #FileNo, JGYOBA_CODE & vbTab & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & vbTab & _
                                               OUT_SYUSI(j) & vbTab & _
                                               Trim(Location1) & vbTab & Trim(Location2) & vbTab & Trim(Location3)
                            Next j
            
                        End If
            
                        Count = Count + 1
                        LblCnt(k).Caption = Format(Count, "#0")
                    
                    
                    End If
        
                    ITEM_com = BtOpGetNext
                Loop
            Next i
        
        
            Close #FileNo
        
            LblFileName(k).Caption = fileName
        
        
        End If
    
    Next k
    
    
    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "ホスト棚番データは正常に出力されました。"

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

    F1200101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200101)


    F1200101.MousePointer = vbDefault

End Sub
                                            'エラーチェック
Private Function Err_Chk() As Integer
                                            
Dim i   As Integer
                                            
                                            
    Err_Chk = True


    For i = ptxYY To ptxDD
    
        If Text(i).Text = "" Then
        
            Select Case i
                Case ptxYY
                    Text(i).Text = "    "
                Case Else
                    Text(i).Text = "  "
            End Select
        
        Else
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(i).SetFocus
                Exit Function
            Else
                If i <> ptxYY Then
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End If
    Next i
    
    Err_Chk = False

End Function

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              'データ出力
            If Err_Chk() Then
                Exit Sub
            End If
        
            Beep
            ans = MsgBox("「ホスト棚番設定データ」出力しますか？", vbYesNo + vbQuestion, "確認入力")
                
            If ans = vbYes Then
                If OUTPUT_Proc Then
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

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String
Dim sts     As Integer

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
    If GetIni("FILE", "HTANA_DATA", "SYS", c) Then
        Beep
        MsgBox "ホスト棚番設定データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    HTANA_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

'    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).Code = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If
'
'        Load SubMenu(i + 1)
'        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
'
'        If JGYOBU_T(i).Code = Last_JGYOBU Then
'            F1200101.Caption = "ホスト棚番設定データ作成（" + RTrim(JGYOBU_T(i).NAME) + ")"
'            SubMenu(i).Checked = True
'            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
''            LabJIGYO.BorderStyle = 1
'        Else
'            SubMenu(i).Checked = False
'        End If
'    Next i
'    Unload SubMenu(i)

'2005/05/31 Add Start #######################################################################
                                        '倉庫最大数を取り込み
    If GetIni(App.EXEName, "MAX_SOKO", "SYS", c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
'2005/05/31 Add End   #######################################################################

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫ﾏｽﾀＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Text(ptxYY).SetFocus
    
    
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
                                            '倉庫ﾏｽﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, ZAIKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫ﾏｽﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1200101 = Nothing

    End
End Sub

'Private Sub SubMenu_Click(Index As Integer)
'Dim i As Integer
'                                    'メニューより終了要求
'    If JGYOBU_T(Index).Code = " " Then
'        Unload Me
'    End If
'
'    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).Code = " " Then
'            Exit For
'        End If
'        SubMenu(i).Checked = False
'    Next i
'                                    '事業部切り替え
'    F1200101.Caption = "ホスト棚番設定データ作成（" + RTrim(JGYOBU_T(Index).NAME) + ")"
'    Last_JGYOBU = JGYOBU_T(Index).Code
'    SubMenu(Index).Checked = True
'
'    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
'    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
'End Sub

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
