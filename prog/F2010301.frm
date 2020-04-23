VERSION 5.00
Begin VB.Form F2010301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品番別在庫データ出力"
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
      Left            =   6120
      MaxLength       =   13
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
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
      MaxLength       =   13
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
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
      Left            =   5640
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
Attribute VB_Name = "F2010301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_HIN_GAI% = 0             '開始　品番
Private Const ptxE_HIN_GAI% = 1             '終了　品番

Private Const Text_Max% = 1                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbNAIGAI% = 0               '国内外

Dim HINZAI_DATA As String                   '品番別在庫データフルパス
'Private Const Last_Update_Day$ = "[F201030]2015.08.20 14:30"
Private Const Last_Update_Day$ = "[F201030]2019.11.06 10:00 品番trim対応"



Private Function OUTPUT_Proc(Mode As Integer) As Integer
    
Dim sts                     As Integer
Dim ZAIKO_com               As Integer
Dim ITEM_com                As Integer

Dim SUMI_ALL_ZAIKO_QTY      As Long
Dim MI_ALL_ZAIKO_QTY        As Long
Dim ALL_ZAIKO_QTY           As Long

Dim SUMI_ST_ZAIKO_QTY       As Long
Dim MI_ST_ZAIKO_QTY         As Long
Dim ST_ZAIKO_QTY            As Long


Dim LOC_ZAIKO_QTY           As Long

Dim SAVE_LOC                As String * 8

Dim Ret                     As Integer
Dim FileNo                  As Integer
Dim FileName                As String

Dim c                       As String * 128
Dim Soko_No                 As String * 2

    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    FileNo = FreeFile
    FileName = HINZAI_DATA
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Error_Proc
    Open (FileName) For Output As FileNo

    Write #FileNo, _
    "品番（外部）", _
    "品名", _
    "総在庫", _
    "標準棚番", _
    "標準棚番" & vbCrLf & "在庫", _
    "別置" & vbCrLf & "棚番1", "別置" & vbCrLf & "在庫1", "別置" & vbCrLf & "棚番2", "別置" & vbCrLf & "在庫2", "別置" & vbCrLf & "棚番3", "別置" & vbCrLf & "在庫3", "別置" & vbCrLf & "棚番4", "別置" & vbCrLf & "在庫4", _
    "別置" & vbCrLf & "棚番5", "別置" & vbCrLf & "在庫5", "別置" & vbCrLf & "棚番6", "別置" & vbCrLf & "在庫6", "別置" & vbCrLf & "棚番7", "別置" & vbCrLf & "在庫7", "別置" & vbCrLf & "棚番8", "別置" & vbCrLf & "在庫8", _
    "別置" & vbCrLf & "棚番9", "別置" & vbCrLf & "在庫9", "別置" & vbCrLf & "棚番10", "別置" & vbCrLf & "在庫10", "別置" & vbCrLf & "棚番11", "別置" & vbCrLf & "在庫11", "別置" & vbCrLf & "棚番12", "別置" & vbCrLf & "在庫12", _
    "別置" & vbCrLf & "棚番13", "別置" & vbCrLf & "在庫13", "別置" & vbCrLf & "棚番14", "別置" & vbCrLf & "在庫14", "別置" & vbCrLf & "棚番15", "別置" & vbCrLf & "在庫15", "別置" & vbCrLf & "棚番16", "別置" & vbCrLf & "在庫16", _
    "別置" & vbCrLf & "棚番17", "別置" & vbCrLf & "在庫17", "別置" & vbCrLf & "棚番18", "別置" & vbCrLf & "在庫18", "別置" & vbCrLf & "棚番19", "別置" & vbCrLf & "在庫19", "別置" & vbCrLf & "棚番20", "別置" & vbCrLf & "在庫20"
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxS_HIN_GAI).Text)

    ITEM_com = BtOpGetGreaterEqual

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

        
        If Zaiko_Syukei_Proc(SUMI_ALL_ZAIKO_QTY, MI_ALL_ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
            Exit Function
        End If

        ALL_ZAIKO_QTY = SUMI_ALL_ZAIKO_QTY + MI_ALL_ZAIKO_QTY

        If Mode = 1 And ALL_ZAIKO_QTY = 0 Then
        Else

            Write #FileNo, Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)),   '2019/11/06 trim対応
            Write #FileNo, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),  '2019/11/06 trim対応

            Write #FileNo, Format(ALL_ZAIKO_QTY, "#0"),

            SAVE_LOC = ""


            If ALL_ZAIKO_QTY = 0 Then
            Else
                                                    '標準棚番分
                If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                    ST_ZAIKO_QTY = 0
                    Write #FileNo, , ,
                Else
                    If Zaiko_Syukei_Proc(SUMI_ST_ZAIKO_QTY, MI_ST_ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                        Exit Function
                    End If
            
                    ST_ZAIKO_QTY = SUMI_ST_ZAIKO_QTY + MI_ST_ZAIKO_QTY
            
                    If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                        Write #FileNo, ,
                    Else
                        If GetIni("SOKO_NO", StrConv(ITEMREC.ST_SOKO, vbUnicode), "SYS", c) Then
                            Soko_No = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                        Else
                            Soko_No = Trim(c)
                        End If
                        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode),
                    End If
                    Write #FileNo, Format(ST_ZAIKO_QTY, "#0"),
                End If
            
                If ALL_ZAIKO_QTY = ST_ZAIKO_QTY Then
                Else
            
                    Call UniCode_Conv(K4_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
                    Call UniCode_Conv(K4_ZAIKO.Retu, "")
                    Call UniCode_Conv(K4_ZAIKO.Ren, "")
                    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
                    LOC_ZAIKO_QTY = 0
                
                    ZAIKO_com = BtOpGetGreater

                    Do
    
                        sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)

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

                        If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) = _
                            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                
                            ZAIKO_com = BtOpGetGreater

                    
                        Else
                            If Len(Trim(SAVE_LOC)) = 0 Then
                                SAVE_LOC = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)

                            End If
                        
                            If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) <> _
                                SAVE_LOC Then
                                
                                If GetIni("SOKO_NO", Left(SAVE_LOC, 2), "SYS", c) Then
                                    Soko_No = Left(SAVE_LOC, 2)
                                Else
                                    Soko_No = Trim(c)
                                End If
                                
                                
                                Write #FileNo, Soko_No & "-" & Mid(SAVE_LOC, 3, 2) & "-" & Mid(SAVE_LOC, 5, 2) & "-" & Right(SAVE_LOC, 2),
                                Write #FileNo, Format(LOC_ZAIKO_QTY, "#0"),
                            
                                SAVE_LOC = (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode))
                                LOC_ZAIKO_QTY = 0
                            End If
                    
                    
                            LOC_ZAIKO_QTY = LOC_ZAIKO_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        
                            ZAIKO_com = BtOpGetNext
                        End If
                    
                
                    Loop
                
                    If Len(Trim(SAVE_LOC)) = 0 Then
                    Else
                        If GetIni("SOKO_NO", Left(SAVE_LOC, 2), "SYS", c) Then
                            Soko_No = Left(SAVE_LOC, 2)
                        Else
                            Soko_No = Trim(c)
                        End If
                        Write #FileNo, Soko_No & "-" & Mid(SAVE_LOC, 3, 2) & "-" & Mid(SAVE_LOC, 5, 2) & "-" & Right(SAVE_LOC, 2),
                        Write #FileNo, Format(LOC_ZAIKO_QTY, "#0"),
            
                    End If
                End If
            End If
        
            Write #FileNo,
        
        
        End If
        
        ITEM_com = BtOpGetNext
    
    Loop


    Close #FileNo
    
    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"

    OUTPUT_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
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

    F2010301.MousePointer = vbHourglass

    Call Ctrl_Lock(F2010301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F2010301)


    F2010301.MousePointer = vbDefault

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
            ans = MsgBox("「品番別在庫データ」出力しますか？", vbYesNo + vbQuestion, "確認入力")
                
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
    If GetIni("FILE", "HINZAI_DATA", "SYS", c) Then
        Beep
        MsgBox "品番別在庫ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    HINZAI_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.20
'    For i = 0 To UBound(JGYOBU_T) - 1
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If
'
'        Load SubMenu(i + 1)
'        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
'
'        If JGYOBU_T(i).CODE = Last_JGYOBU Then
'            F2010301.Caption = "品番別在庫データ出力（" + RTrim(JGYOBU_T(i).NAME) + ")"
'            SubMenu(i).Checked = True
'            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
''            LabJIGYO.BorderStyle = 1
'        Else
'            SubMenu(i).Checked = False
'        End If
'    Next i


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F2010301.Caption = "品番別在庫データ出力（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_Day

            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.20


                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
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
    
    Set F2010301 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)

'>>>>>>>>>>>>>>>>>>>>>> 2015.08.20
'Dim i As Integer
'                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
'
'    For i = 0 To UBound(JGYOBU_T) - 1
'        If JGYOBU_T(i).CODE = " " Then
'            Exit For
'        End If
'        SubMenu(i).Checked = False
'    Next i
'                                    '事業部切り替え
'    F2010301.Caption = "品番別在庫データ出力（" + RTrim(JGYOBU_T(Index).NAME) + ")"
'    Last_JGYOBU = JGYOBU_T(Index).CODE
'    SubMenu(Index).Checked = True
'
'    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
'    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)



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
    F2010301.Caption = "品番別在庫一覧表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")" & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
'>>>>>>>>>>>>>>>>>>>>>> 2015.08.20

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


