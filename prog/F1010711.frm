VERSION 5.00
Begin VB.Form F1010711 
   BackColor       =   &H00FFFFFF&
   Caption         =   "[作業管理マスタ]メニュー管理メンテナンス"
   ClientHeight    =   9315
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   13740
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
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13740
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   120
      MaxLength       =   2
      TabIndex        =   49
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   48
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   9
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   47
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "=>"
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   46
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "=>"
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   45
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行削除"
      Height          =   375
      Index           =   5
      Left            =   12480
      TabIndex        =   44
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行削除"
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   43
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行削除"
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   42
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行設定"
      Height          =   375
      Index           =   2
      Left            =   12480
      TabIndex        =   41
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行設定"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   40
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "主要因"
      Height          =   4455
      Left            =   1560
      TabIndex        =   37
      Top             =   2640
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         Top             =   3960
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   3420
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.ListBox List1 
      Height          =   5820
      Index           =   2
      Left            =   8760
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   12
      Left            =   11040
      TabIndex        =   34
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   11
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   33
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   10
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   32
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "行設定"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   30
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   28
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   26
      Top             =   1680
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   5820
      Index           =   1
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   23
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   960
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   480
      MaxLength       =   1
      TabIndex        =   17
      Top             =   1680
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   5820
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.ComboBox cboNAIGAI 
      Height          =   360
      ItemData        =   "F1010711.frx":0000
      Left            =   2400
      List            =   "F1010711.frx":0007
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox cboJIGYOBU 
      Height          =   360
      ItemData        =   "F1010711.frx":0019
      Left            =   240
      List            =   "F1010711.frx":0020
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   1  'ｵﾝ
      Index           =   1
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   3
      Top             =   360
      Width           =   2535
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8520
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
      Index           =   7
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ｇ削除"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label lbMenu_LV2 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   8040
      TabIndex        =   51
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lbMenu_LV1 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   3120
      TabIndex        =   50
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "表示名称"
      Height          =   375
      Index           =   7
      Left            =   11040
      TabIndex        =   35
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "向け先"
      Height          =   255
      Index           =   6
      Left            =   8880
      TabIndex        =   31
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "表示名称"
      Height          =   375
      Index           =   10
      Left            =   5880
      TabIndex        =   29
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "要因"
      Height          =   375
      Index           =   9
      Left            =   5400
      TabIndex        =   27
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "表示名称"
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "要因"
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "区分"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "区分:0:メニュー 1:作業"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "F1010711"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NAIGAI_CODE()   As String * 1

Private Yoin_Tbl()      As String * 1


Private ALL_MENU_GRP    As String * 2

Private Const ETC_CODE_TYPE$ = "?"
Private Const ETC_CODE_TYPE_NAME$ = "ＥＴＣ"


Private Sub Command_Click(Index As Integer)

    Select Case Index
        Case 11
            Unload Me
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)

Dim i                           As Integer
Dim Err_Flg                     As Boolean

Dim Edit                        As String
Dim sts                         As Integer

Dim ED_CNT                      As String * 2
Dim ED_MENU_KBN                 As String * 1
'Dim ED_CODE_TYPE                As String * 1
Dim ED_YOIN                     As String * 2
Dim ED_DISPLAY_ITEM(0 To 19)    As Byte
Dim ED_LV                       As String * 3

Dim New_Flg     As Integer

    Select Case Index
        
        Case 0
        
            If Text(3).Text <> "0" And Text(3).Text <> "1" Then
                MsgBox "区分エラー"
                Text(3).SetFocus
                Exit Sub
            End If
        
            Select Case Text(3).Text
                
                Case "0"
            
                    Err_Flg = True
            
            
                    If Text(4).Text <> ETC_CODE_TYPE Then
            
                        For i = 0 To UBound(Yoin_Tbl)
                            If Text(4).Text = Yoin_Tbl(i) Then
                                Err_Flg = False
                                Exit For
                            End If
                        Next i
                        
                        If Err_Flg Then
                            MsgBox "主CDエラー"
                            Text(4).SetFocus
                            Exit Sub
                        End If
                    End If
                
                Case "1"
            
            
                    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Text(4).Text, 1))
                    Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Text(4).Text, 1))
            
                    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                    Select Case sts
                        Case BtNoErr
                        
'                            Text(6).Text = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
                        Case BtErrKeyNotFound
                            Text(6).Text = ""
                            MsgBox "入力した項目はエラーです。"
                            Text(5).SetFocus
                            Exit Sub
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                            Exit Sub
                    End Select
            
            
            End Select
            
            For i = 0 To List1(0).ListCount - 1
            
                If Left(List1(0).List(i), 2) = Text(2).Text Then
                    List1(0).RemoveItem i
                    Exit For
                End If
            
            Next i

            New_Flg = False
            If Not IsNumeric(lbMenu_LV1.Caption) Then
                lbMenu_LV1.Caption = Format(List1(0).ListCount, "000")
                New_Flg = True
            End If



            ED_CNT = Text(2).Text
            ED_MENU_KBN = Text(3).Text
'            ED_CODE_TYPE = Text(4).Text
            ED_YOIN = Text(4).Text
            Call UniCode_Conv(ED_DISPLAY_ITEM, Text(5).Text)
            ED_LV = lbMenu_LV1.Caption
            
            Edit = ED_CNT & ED_MENU_KBN & ED_CODE_TYPE & ED_YOIN_CODE & StrConv(ED_DISPLAY_ITEM, vbUnicode) & ED_LV
                                    
            List1(0).AddItem Edit
            
            
            Select Case New_Flg
                Case False
                    Do
                        
                        
                                                
                        
                        
                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(ED_YOIN_CODE, 1))
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(ED_YOIN_CODE, 1))
                        
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                        
'                                Text(6).Text = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
                            Case BtErrKeyNotFound
                                Text(6).Text = ""
                                MsgBox "入力した項目はエラーです。"
                                Text(3).SetFocus
                                Exit Sub
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                                Exit Sub
                        End Select
                
            
                    Loop
            
            
            
                Case True
            
            End Select
            
            
        
        
        Case 3
        
            For i = 0 To List1(0).ListCount - 1
            
                If Left(List1(0).List(i), 2) = Text(2).Text Then
                    List1(0).RemoveItem i
                    Exit For
                End If
            
            Next i
        
        
        Case 6
    
            If Len(Trim(lbMenu_LV1.Caption)) = 0 Then
                Exit Sub
            End If
    
    
            If List_Disp_Proc(1) Then
                Unload Me
            End If
    
        Case 7
    
            If Len(Trim(lbMenu_LV2.Caption)) = 0 Then
                Exit Sub
            End If
    
    
            If List_Disp_Proc(2) Then
                Unload Me
            End If
    
    
    End Select


End Sub

Private Sub Command2_Click()

    Frame1.Visible = False
    Text(3).SetFocus
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
    
Dim Edit    As String
    
Dim YOIN    As String * 1
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
    If JGYOB_TB_Set() Then
        Beep
        MsgBox "事業部の獲得に失敗しました。"
        End
    End If


    cboJIGYOBU.Clear
    For i = 0 To UBound(JGYOBU_T)
        If Trim(JGYOBU_T(i).Code) = "" Then
            Exit For
        End If
        cboJIGYOBU.AddItem JGYOBU_T(i).NAME & " " & JGYOBU_T(i).Code
    Next i
    cboJIGYOBU.ListIndex = 0
    If cboJIGYOBU.ListCount = 1 Then
        cboJIGYOBU.Enabled = False
    End If
    
    
    '国内外情報設定
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        Beep
        MsgBox "国内外の獲得に失敗しました。"
        End
    End If

    cboNAIGAI.Clear
    For i = 0 To UBound(NAIGAI_CODE)
        
        Select Case NAIGAI_CODE(i)
            Case NAIGAI_NAI
                cboNAIGAI.AddItem NAIGAI1 & " " & NAIGAI_CODE(i)
        
            Case NAIGAI_GAI
                cboNAIGAI.AddItem NAIGAI2 & " " & NAIGAI_CODE(i)
        End Select
                    
    Next i
    cboNAIGAI.ListIndex = 0
    If cboNAIGAI.ListCount = 1 Then
        cboNAIGAI.Enabled = False
    End If

                                'メニューマスタＯＰＥＮ
    If MENU_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If


    List2.Clear
    
        '作業タイプ設定
    i = 0
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
        Edit = Trim(c)
        YOIN = Trim(c)
    
        If GetIni("ACTION", "ACTION_TYPE" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[F101070]" & "[ACTION_TYPE" & Format(i, "00") & "]"
            Exit Do
        End If
        If Trim(c) = "1" Then
        Else
'            Edit = Trim(c)
        
            ReDim Preserve Yoin_Tbl(i)
            Yoin_Tbl(i) = YOIN
            
            
            If GetIni("ACTION", "ACTION_NM" & Format(i + 1, "00"), "SYS", c) Then
                MsgBox "作業情報の獲得に失敗しました。" & "[F101070]" & "[ACTION_NM" & Format(i, "00") & "]"
                Exit Do
            End If
            Edit = Edit & Trim(c)
            List2.AddItem Edit
        
        End If
        i = i + 1
    
    Loop

    List2.AddItem ETC_CODE_TYPE & ETC_CODE_TYPE_NAME
    


    Show

    Text(0).SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
                                            
Dim sts As Integer
                                            'メニュー管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ")
        End If
    End If

                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If

                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010711 = Nothing
    End

End Sub


Private Sub List1_DblClick(Index As Integer)

    Select Case Index

        Case 0
            Text(2).Text = Left(List1(Index).List(List1(Index).ListIndex), 2)
            Text(3).Text = Mid(List1(Index).List(List1(Index).ListIndex), 3, 1)
            Text(4).Text = Mid(List1(Index).List(List1(Index).ListIndex), 4, 2)
            Text(5).Text = Mid(List1(Index).List(List1(Index).ListIndex), 6, Len(List1(Index).List(List1(Index).ListIndex)) - 6 - 3)
            lbMenu_LV1.Caption = Right(List1(Index).List(List1(Index).ListIndex), 3)
    
    
            Text(2).SetFocus
    
    
        Case 1
    
            Text(7).Text = Left(List1(Index).List(List1(Index).ListIndex), 2)
            Text(8).Text = Mid(List1(Index).List(List1(Index).ListIndex), 3, 2)
            Text(9).Text = Mid(List1(Index).List(List1(Index).ListIndex), 6, Len(List1(Index).List(List1(Index).ListIndex)) - 4 - 3)
            lbMenu_LV2.Caption = Right(List1(Index).List(List1(Index).ListIndex), 3)
    
    
            Text(7).SetFocus
    
    End Select

End Sub

Private Sub List2_Click()

    Text(4).Text = Left(List2.List(List2.ListIndex), 1)
    Text(6).Text = Right(List2.List(List2.ListIndex), Len(List2.List(List2.ListIndex)) - 1)

    Frame1.Visible = False

    Text(6).SetFocus
End Sub


Private Sub Text_DblClick(Index As Integer)
    If Index = 4 Then
        Frame1.Visible = True
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i       As Integer
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        
        Case 0          'メニューグループ番号
            If Len(Trim(Text(0).Text)) = 0 Then
                Beep
                MsgBox "入力した項目はエラーです。（必須入力）"
                Text(0).SetFocus
                Exit Sub
            End If
                        
                                                
            Text(1).SetFocus
                        
                                    
                        
                        
            If List_Disp_Proc(0) Then
                Unload Me
            End If
                        
                        
    
        Case 1          'メニュー名称
        
        
        
        Case 3          '区分
            If Text(3).Text <> "0" And Text(3).Text <> "1" Then
                Beep
                MsgBox "入力した項目はエラーです。（0 OR 1）"
                Text(3).SetFocus
                Exit Sub
            End If
    
            
            
            Select Case Text(3).Text
                Case "0"
                
                    Text(4).MaxLength = 1
                
                Case "1"
                    Text(4).MaxLength = 2
            
            End Select
    
    
                
    
    
    
        Case 4          '主CD&要因
    
        
            If Text(3).Text = "0" Then
                For i = 0 To List2.ListCount - 1
                
                    If Text(4).Text = Left(List2.List(i), 1) Then
                        Text(5).Text = Right(List2.List(i), Len(List2.List(i)) - 1)
                        Exit For
                    End If
                
                Next i
            
            
                If i > (List2.ListCount - 1) Then
                    Text(5).Text = ""
                    MsgBox "入力した項目はエラーです。"
                    Text(4).SetFocus
                    Exit Sub
                End If
            
            Else
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Text(4).Text, 1))
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Text(4).Text, 1))
        
                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                        
                        Text(5).Text = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
                    
                    
                        If StrConv(YOINREC.REGI_F, vbUnicode) = "0" Or _
                            StrConv(YOINREC.REGI_F, vbUnicode) = "1" Then
                        Else
                            MsgBox "スキャナメニューヘの登録は出来ません。"
                            Text(4).SetFocus
                            Exit Sub
                        End If
                    
                    Case BtErrKeyNotFound
                        Text(5).Text = ""
                        MsgBox "入力した項目はエラーです。"
                        Text(4).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                        Exit Sub
                End Select
            End If
    
    
        Case 5          '名称
    
    
    
    End Select

    For i = Index + 1 To 12
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i


End Sub

Public Function List_Disp_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   画面表示処理
'----------------------------------------------------------------------------

Dim com     As Integer
Dim sts     As Integer

Dim Edit    As String

Dim H_Cnt   As Integer


    List_Disp_Proc = True

    Select Case Mode
    
        Case 0
            List1(Mode).Clear
    
            com = BtOpGetGreater
    
            Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(0).Text)
            Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU, 1))
            Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI, 1))
    
            Call UniCode_Conv(K0_MENU.MENU_LV1, "")
            Call UniCode_Conv(K0_MENU.MENU_LV2, "")
            Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
            H_Cnt = 0
    
            Do
                sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                Select Case sts
                    Case BtNoErr
                    
                        If Text(0).Text <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Or _
                            Right(cboJIGYOBU, 1) <> StrConv(MENUREC.JGYOBU, vbUnicode) Or _
                            Right(cboNAIGAI, 1) <> StrConv(MENUREC.NAIGAI, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                        Exit Function
                End Select
            
                If Len(Trim(StrConv(MENUREC.MENU_LV2, vbUnicode))) <> 0 Or _
                    Len(Trim(StrConv(MENUREC.MENU_LV3, vbUnicode))) <> 0 Then
                Else
                                
                    Text(1).Text = StrConv(MENUREC.MENU_GRP, vbUnicode)
                    
                    H_Cnt = H_Cnt + 3
                    Edit = Format(H_Cnt, "00")
                    Edit = Edit & StrConv(MENUREC.MENU_KBN, vbUnicode)
                    Edit = Edit & StrConv(MENUREC.CODE_TYPE, vbUnicode)
                    Edit = Edit & StrConv(MENUREC.YOIN_CODE, vbUnicode)
                    
                    Edit = Edit & StrConv(MENUREC.DISPLAY_ITEM, vbUnicode) & StrConv(MENUREC.MENU_LV1, vbUnicode) & lbMenu_LV1.Caption
                                
                                
                    List1(Mode).AddItem Edit
                End If
                
                
                
                
            Loop
    
    
        Case 1
            List1(Mode).Clear
    
            com = BtOpGetGreater
    
            Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(0).Text)
            Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU, 1))
            Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI, 1))
    
            Call UniCode_Conv(K0_MENU.MENU_LV1, lbMenu_LV1)
            Call UniCode_Conv(K0_MENU.MENU_LV2, "")
            Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
            H_Cnt = 0
    
            Do
                sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                Select Case sts
                    Case BtNoErr
                    
                        If Text(0).Text <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Or _
                            Right(cboJIGYOBU, 1) <> StrConv(MENUREC.JGYOBU, vbUnicode) Or _
                            Right(cboNAIGAI, 1) <> StrConv(MENUREC.NAIGAI, vbUnicode) Then
                            Exit Do
                        End If
                    
                        If lbMenu_LV1 <> StrConv(MENUREC.MENU_LV1, vbUnicode) Then
                            Exit Do
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                        Exit Function
                End Select
            
                If Len(Trim(StrConv(MENUREC.MENU_LV3, vbUnicode))) <> 0 Then
                Else
                                
                    Text(1).Text = StrConv(MENUREC.MENU_GRP, vbUnicode)
                    
                    H_Cnt = H_Cnt + 1
                    Edit = Format(H_Cnt, "00")
                    Edit = Edit & StrConv(MENUREC.MENU_KBN, vbUnicode)
                    Edit = Edit & StrConv(MENUREC.CODE_TYPE, vbUnicode)
                    Edit = Edit & StrConv(MENUREC.YOIN_CODE, vbUnicode)
                    Edit = Edit & StrConv(MENUREC.DISPLAY_ITEM, vbUnicode) & StrConv(MENUREC.MENU_LV2, vbUnicode)
                                
                                
                    List1(Mode).AddItem Edit
                End If
                
            Loop
    
    
        Case 2
            List1(Mode).Clear
    
            com = BtOpGetGreater
    
            Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(0).Text)
            Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU, 1))
            Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI, 1))
    
            Call UniCode_Conv(K0_MENU.MENU_LV1, lbMenu_LV1)
            Call UniCode_Conv(K0_MENU.MENU_LV2, lbMenu_LV2)
            Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
            H_Cnt = 0
    
            Do
                sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                Select Case sts
                    Case BtNoErr
                    
                        If Text(0).Text <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Then
                            Exit Do
                        End If
                    
                        If lbMenu_LV1 <> StrConv(MENUREC.MENU_LV1, vbUnicode) Or _
                            lbMenu_LV2 <> StrConv(MENUREC.MENU_LV2, vbUnicode) Then
                            Exit Do
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                        Exit Function
                End Select
            
                                
                Text(1).Text = StrConv(MENUREC.MENU_GRP, vbUnicode)
                
                H_Cnt = H_Cnt + 1
                Edit = Format(H_Cnt, "00")
                Edit = Edit & StrConv(MENUREC.CODE_TYPE, vbUnicode)
                Edit = Edit & StrConv(MENUREC.YOIN_CODE, vbUnicode)
                Edit = Edit & StrConv(MENUREC.DISPLAY_ITEM, vbUnicode) & StrConv(MENUREC.MENU_LV2, vbUnicode)
                            
                            
                List1(Mode).AddItem Edit
                
            Loop
    
    End Select


    List_Disp_Proc = False

End Function
