VERSION 5.00
Begin VB.Form F1060211 
   BackColor       =   &H00FFFFFF&
   Caption         =   "「商品化実績対応」商品化計画支援アラームリスト印刷 "
   ClientHeight    =   6948
   ClientLeft      =   2328
   ClientTop       =   2712
   ClientWidth     =   11292
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
   ScaleHeight     =   6948
   ScaleWidth      =   11292
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   8715
      MaxLength       =   2
      TabIndex        =   34
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   8085
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   30
      Top             =   3600
      Width           =   645
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   6300
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   5670
      MaxLength       =   2
      TabIndex        =   26
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   4620
      MaxLength       =   4
      TabIndex        =   24
      Top             =   3600
      Width           =   645
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   4620
      TabIndex        =   21
      Top             =   2760
      Width           =   330
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   4620
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   2160
      Width           =   1170
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   5460
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   17
      Top             =   1440
      Width           =   3270
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4620
      TabIndex        =   16
      Top             =   1440
      Width           =   750
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4575
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   840
      Width           =   1125
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
         Size            =   11.4
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
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   11
      Left            =   9030
      TabIndex        =   35
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   10
      Left            =   8400
      TabIndex        =   33
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   9
      Left            =   7770
      TabIndex        =   31
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日～"
      Height          =   255
      Index           =   8
      Left            =   6615
      TabIndex        =   29
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   7
      Left            =   6090
      TabIndex        =   27
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   6
      Left            =   5355
      TabIndex        =   25
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "対象年月日"
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   23
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（※空白　全倉庫指定）"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   22
      Top             =   2880
      Width           =   2730
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標準棚番（倉庫番号）"
      Height          =   255
      Index           =   3
      Left            =   1995
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "未商品在庫"
      Height          =   255
      Index           =   2
      Left            =   3150
      TabIndex        =   18
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "受払先コード"
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "取引先区分"
      Height          =   255
      Index           =   0
      Left            =   3210
      TabIndex        =   14
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   12
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
Attribute VB_Name = "F1060211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxUKEHARAI_CODE% = 0         '受払先コード
Private Const ptxSOKO_NO% = 1               '倉庫番号
Private Const ptxS_YY% = 2                  '開始年月日　年
Private Const ptxS_MM% = 3                  '開始年月日　月
Private Const ptxS_DD% = 4                  '開始年月日　日
Private Const ptxE_YY% = 5                  '終了年月日　年
Private Const ptxE_MM% = 6                  '終了年月日　月
Private Const ptxE_DD% = 7                  '終了年月日　日


Private Const Text_Max% = 7                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbTORI_KBN% = 0             '取引先コード
Private Const pcmbUKEHARAI_CODE% = 1        '受払先コード
Private Const pcmbMI_ZAIKO% = 2             '未商品在庫


Private Const LMAX% = 36                    '頁内最大行数
Private Const LCTL% = 99                    '
Private Const MGN_L% = 10                   '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private NormalFont  As New StdFont          '印刷フォント
Private MidFont     As New StdFont          '印刷フォント

Private OutSide     As Long                 '印刷対外出荷数

Private GOODS_DATA  As String               '出力データファイル名

Private NON_MTS     As String               '除外向け先


Private Type EE_ZAIKO_TBL_tag
    EE_LOC          As String * 8
    EE_QTY          As Long
End Type

Private EE_ZAIKO_TBL(0 To 2) As EE_ZAIKO_TBL_tag

Private SHO_SOKO    As Variant              '商品化用倉庫(未商品とみなす分)

Private Const Last_Update_day$ = "([F106021] 2011.07.14 12:00)"

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   エラーチェック処理
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
            
            
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060211.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060211)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060211)


    F1060211.MousePointer = vbDefault

End Sub



Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)


    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
    
        Case pcmbTORI_KBN
        
        
        
            If Ukeharai_Set_Proc() Then
                Unload Me
            End If
        
            Combo(pcmbUKEHARAI_CODE).SetFocus
        
        
        
        Case pcmbUKEHARAI_CODE
        
            Text(ptxUKEHARAI_CODE).Text = Right(Combo(pcmbUKEHARAI_CODE).Text, 5)
            Combo(pcmbMI_ZAIKO).SetFocus
        
        
        Case pcmbMI_ZAIKO
    
            Text(ptxSOKO_NO).SetFocus
    
    End Select


End Sub

Private Sub Combo_LostFocus(Index As Integer)

Dim i   As Integer


    Select Case Index


        Case pcmbTORI_KBN



            If Ukeharai_Set_Proc() Then
                Unload Me
            End If

'            Combo(pcmbUKEHARAI_CODE).SetFocus

            If Trim(Text(ptxUKEHARAI_CODE).Text) <> "" Then
                For i = 0 To Combo(pcmbUKEHARAI_CODE).ListCount - 1
                    If Trim(Text(ptxUKEHARAI_CODE).Text) = Trim(Right(Combo(pcmbUKEHARAI_CODE).List(i), 5)) Then
                        Combo(pcmbUKEHARAI_CODE).ListIndex = i
                        Exit For
                    End If
                Next i
            End If




        Case pcmbUKEHARAI_CODE

            Text(ptxUKEHARAI_CODE).Text = Right(Combo(pcmbUKEHARAI_CODE).Text, 5)
'            Combo(pcmbMI_ZAIKO).SetFocus


        Case pcmbMI_ZAIKO

'            Text(ptxSOKO_NO).SetFocus
    
    End Select



End Sub

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
    Select Case Index
        
        Case 7                              'データ出力
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                If Data_Proc() Then
                    Unload Me
                End If
            End If
            
            Text(ptxS_YY).SetFocus
        
        
        Case 8                              '印刷
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxS_YY).SetFocus
                    
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
            F1060211.Caption = "商品化計画支援アラームリスト(出荷ﾃﾞｰﾀ対応)印刷（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '商品化支援ファイル名取り込み
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'商品化支援ファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    GOODS_DATA = Trim(c)
                                
'-----------    SYS.INI ---> (自ﾌﾟﾛｸﾞﾗﾑID).INI 2011.07.14
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
                                '商品化用倉庫取り込み
    If GetIni(App.EXEName, "SHO_SOKO", "SYS", c) Then
        c = " "
    End If
    SHO_SOKO = Split(Trim(c), ",", -1)
                                
'-----------    SYS.INI ---> (自ﾌﾟﾛｸﾞﾗﾑID).INI 2011.07.14
                                
                                
                                '除外向け先取り込み
    If GetIni("PI00010", "MTSSS", "P_SYS", c) Then
        NON_MTS = ""
    Else
        NON_MTS = Trim(c)
    End If
                                
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
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
                                '商品化集計ファイルＯＰＥＮ
    If GOODS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図データ（親）ＯＰＥＮ 2007.11.14
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定(通常)
    With NormalFont
        .NAME = F1060211.FontName
        .Size = 12
    End With

                                '印刷フォント設定（小）
    With MidFont
        .NAME = F1060211.FontName
        .Size = 8
    End With





    Combo(pcmbTORI_KBN).AddItem "全　て　     " & " "
    Combo(pcmbTORI_KBN).AddItem P_TORI_GENERAL_N & "     " & P_TORI_GENERAL
    Combo(pcmbTORI_KBN).AddItem P_TORI_NAISYOKU_N & "     " & P_TORI_NAISYOKU
    Combo(pcmbTORI_KBN).AddItem P_TORI_GENKIN_N & "     " & P_TORI_GENKIN
    Combo(pcmbTORI_KBN).AddItem P_TORI_SYANAI_N & "     " & P_TORI_SYANAI
    Combo(pcmbTORI_KBN).AddItem P_TORI_ANOTHER_N & "     " & P_TORI_ANOTHER
    Combo(pcmbTORI_KBN).AddItem P_TORI_JIKYU_N & "     " & P_TORI_JIKYU
    Combo(pcmbTORI_KBN).ListIndex = 0

    Combo(pcmbMI_ZAIKO).AddItem "全　て　     " & "0"
    Combo(pcmbMI_ZAIKO).AddItem "０除く　     " & "1"
    Combo(pcmbMI_ZAIKO).AddItem "０のみ　     " & "2"
    Combo(pcmbMI_ZAIKO).ListIndex = 0


    

    Show
    
    Combo(pcmbTORI_KBN).SetFocus
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
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
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
                                            '商品化指図データ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図データ(親)")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060211 = Nothing

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
    F1060211.Caption = "商品化計画支援アラームリスト(出荷ﾃﾞｰﾀ対応)印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）" & Last_Update_day
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

    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
        
        Case ptxUKEHARAI_CODE
        
            For i = 0 To Combo(pcmbUKEHARAI_CODE).ListCount - 1
                If Trim(Text(ptxUKEHARAI_CODE).Text) = Trim(Right(Combo(pcmbUKEHARAI_CODE).List(i), 5)) Then
                    Combo(pcmbUKEHARAI_CODE).ListIndex = i
                    Exit For
                End If
            Next i
        
            If i > Combo(pcmbUKEHARAI_CODE).ListCount - 1 Then
        
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(Index).SetFocus
                Exit Sub
        
            End If
        
        Case ptxSOKO_NO
        
        Case ptxS_YY
        
                    
        Case ptxS_MM, ptxS_DD
    
            If Trim(Text(Index).Text) = "" Then
            Else
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
        Case ptxE_YY
        
                    
        Case ptxE_MM, ptxE_DD
    
            
            
            If Trim(Text(Index).Text) = "" Then
            Else
            
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
        
            If Index = ptxE_DD Then
                If Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text > _
                    Text(ptxE_YY).Text & Text(ptxE_MM).Text & Text(ptxE_DD).Text Then
    
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Text(ptxS_YY).SetFocus
                    Exit Sub
                End If
            End If
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
Dim LCNT        As Integer

Dim sts         As Integer
Dim com         As Integer
Dim FSW         As Boolean

Dim Save_Soko   As String * 2

Dim Edit        As String

Dim SKIP_Flg    As Boolean

Dim X_Tab       As Integer

    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    End If



    LCNT = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    Call UniCode_Conv(K1_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K1_GOODS.ST_RETU, "")
    Call UniCode_Conv(K1_GOODS.ST_REN, "")
    Call UniCode_Conv(K1_GOODS.ST_DAN, "")
    Call UniCode_Conv(K1_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K1_GOODS.HIN_GAI, "")
    
    
    com = BtOpGetGreater
    FSW = True
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select


'-------------------------------------------------  明細印刷
            
        SKIP_Flg = False
        Select Case Right(Combo(pcmbMI_ZAIKO).Text, 1)
            Case "0"        '全て
            Case "1"        '0対象外
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) = 0 Then
                    SKIP_Flg = True
                End If
            Case "2"        '0のみ
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <> 0 Then
                    SKIP_Flg = True
                End If
        End Select
            
            
            
            
        
        If SKIP_Flg Then
        Else
            
            If FSW Then
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
                FSW = False
            End If
            
            
            
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                LCNT = LMAX + 1
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
            
            
            
            
            
            If Head_Print_Proc(LCNT) Then
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

            Printer.Print Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 5
            X_Tab = X_Tab + Len(Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13)) + 4
                                                    '箱№
            Printer.Print Tab(X_Tab);
            Printer.Print StrConv(GOODSREC.PACKING_NO, vbUnicode);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 5
            X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
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
            Edit = Format(CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If

            Printer.Print Edit;
            X_Tab = X_Tab + Len(Edit) + 5
                                                    '別置在庫
            Printer.Print Tab(X_Tab);

            If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If

            Edit = ""
            If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
                Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#,##0")
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
        
            LCNT = LCNT + 2
    
        End If
        com = BtOpGetNext
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(LCNT As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If LCNT < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    If LCNT = LCTL Then
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

    Printer.Print Tab(MGN_L + 35);
    
    Printer.Print "商品化支援アラームリスト(出荷データ対応)";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "倉庫：";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  "
'    Printer.Print "（設定発注点 " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "％）"
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
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "資材";
    Printer.Print Tab(MGN_L + 42);
    Printer.Print "商済数";
    Printer.Print Tab(MGN_L + 54);
    Printer.Print "未商品";
    Printer.Print Tab(MGN_L + 62);
    Printer.Print "本日出荷数";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "必要数";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "　状況";
    Printer.Print Tab(MGN_L + 113);
    Printer.Print "別置在庫"

    Printer.Print

    LCNT = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   支援用集計データ作成処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer


Dim SKIP_Flg    As Boolean
    
    Data_Make_Proc = True

'---------------------------------------------------------- '全レコード削除
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
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
'---------------------------------------------------------- '指図票データベースで作成（KEYのみ）

    
    
    Call UniCode_Conv(K3_P_SSHIJI_O.HAKKO_DT, Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text)
    Call UniCode_Conv(K3_P_SSHIJI_O.TORI_KBN, Right(Combo(pcmbTORI_KBN).Text, 1))
    Call UniCode_Conv(K3_P_SSHIJI_O.UKEHARAI_CODE, Text(ptxUKEHARAI_CODE).Text)
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K3_P_SSHIJI_O, Len(K3_P_SSHIJI_O), 3)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode) > Text(ptxE_YY).Text & _
                                                                    Text(ptxE_MM).Text & _
                                                                    Text(ptxE_DD).Text Then
                    '日付範囲外
                    Exit Do
                End If
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化指図表データ")
                Exit Function
        End Select
        
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
        SKIP_Flg = False
        
        
        If Trim(Right(Combo(pcmbTORI_KBN).Text, 1)) <> "" Then
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> Right(Combo(pcmbTORI_KBN).Text, 1) Then
                '取引先区分
                SKIP_Flg = True
            End If
        End If
        
        
        
        
        If Trim(Text(ptxUKEHARAI_CODE).Text) <> "" Then
            If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text(ptxUKEHARAI_CODE).Text) Then
                '取引先コード（受払先コード）
                SKIP_Flg = True
            End If
        End If
        
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
        
        If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
            StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
            SKIP_Flg = True
        End If
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> GOODS_ON Then
                    SKIP_Flg = True
                End If
            Case BtErrKeyNotFound
                SKIP_Flg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
                
                
        If Trim(Text(ptxSOKO_NO).Text) <> "" Then
            If Text(ptxSOKO_NO).Text <> StrConv(ITEMREC.ST_SOKO, vbUnicode) Then
                SKIP_Flg = True
            End If
        End If
            
            
            
        If Not SKIP_Flg Then
            
            Call UniCode_Conv(K2_GOODS.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K2_GOODS.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K2_GOODS.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            
            sts = BTRV(BtOpGetEqual, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
            Select Case sts
                Case BtNoErr
                    SKIP_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化集計ファイル")
                    Exit Function
            End Select
            
            
            If Not SKIP_Flg Then
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
                                                        '事業部
                Call UniCode_Conv(GOODSREC.JGYOBU, Last_JGYOBU)
                                                        '国内外
                Call UniCode_Conv(GOODSREC.NAIGAI, NAIGAI_NAI)
                                                        '品番（外部）
                Call UniCode_Conv(GOODSREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                        '標準棚番
                Call UniCode_Conv(GOODSREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                        '箱№
                Call UniCode_Conv(GOODSREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
            
            
                Call UniCode_Conv(GOODSREC.Sumi_QTY, "00000000")        '商品化済み在庫数
                Call UniCode_Conv(GOODSREC.Mi_QTY, "00000000")          '未商品在庫数
                Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")       '平均出荷数
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")    '事前商品化状況
            
            
                sts = BTRV(BtOpInsert, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "商品化計画支援")
                        Exit Function
                End Select
            
            End If
            
        End If
            
        com = BtOpGetNext
            
    Loop
            
            
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "商品化集計ファイル")
                Exit Function
        End Select
Debug.Print StrConv(GOODSREC.HIN_GAI, vbUnicode)
        If Data_Make_Sub() Then
            Exit Function
        End If
            
        
        com = BtOpGetNext
    
    Loop


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
Dim fileName        As String
    
    
Dim SKIP_Flg        As Boolean
    
    
    Data_Proc = True

    Call Input_Lock

    fileName = GOODS_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - sts)
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo
    On Error GoTo 0


    If Data_Make_Proc() Then        '商品化支援集計データ作成
        Exit Function
    End If
    
    Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GOODS.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化集計ファイル")
                Exit Function
        End Select
'-------------------------------------------------  明細印刷
        
        SKIP_Flg = False
        Select Case Right(Combo(pcmbMI_ZAIKO).Text, 1)
            Case "0"        '全て
            Case "1"        '0対象外
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) = 0 Then
                    SKIP_Flg = True
                End If
            Case "2"        '0のみ
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <> 0 Then
                    SKIP_Flg = True
                End If
        End Select
        
        If SKIP_Flg Then
        Else
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
                        'ヘッダー出力
                Write #FileNo, "*** 商品化支援アラームリスト　***",
                Write #FileNo, "作成日付:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                        
            
                Write #FileNo, "標準棚番", "品番（外部）", "資材（箱№）", "商品化済在庫", "未商品在庫", "未商品　別置き1", "未商品　別置き2", "未商品　別置き3", "月平均出荷数", "事前商品化必要数", "事前商品化状況"
                
            
                Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
            
            
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
                
                Write #FileNo, "倉庫№：" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(発注点" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
                
            End If
        
        
            
            
                                                    '標準棚番
                            
            Edit = StrConv(SOKOREC.Soko_No, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
            Write #FileNo, Edit,
                                                    '品番（外部）

            Write #FileNo, StrConv(GOODSREC.HIN_GAI, vbUnicode),
                                                    '箱№
            Write #FileNo, StrConv(GOODSREC.PACKING_NO, vbUnicode),
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
                                                    
            If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If
                                                    '未商品別置き
            If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(0).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
            If Len(Trim(EE_ZAIKO_TBL(1).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(1).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(1).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
            If Len(Trim(EE_ZAIKO_TBL(2).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(2).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(2).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(2).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
                                                    '月平均出荷数
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '事前商品化必要数
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '事前商品化状況
            Edit = Format(CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit
                
            com = BtOpGetNext
        End If
    Loop

    Close #FileNo

    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    Call Input_UnLock
    
    Data_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Proc = True
    End If


End Function

Private Function MI_ZAIKO_KENSAKU(Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   未商品の処理
'----------------------------------------------------------------------------
Dim i           As Integer

Dim com         As Integer
Dim sts         As Integer

    MI_ZAIKO_KENSAKU = True
    
    For i = 0 To UBound(EE_ZAIKO_TBL)
        EE_ZAIKO_TBL(i).EE_LOC = ""
        EE_ZAIKO_TBL(i).EE_QTY = 0
    Next i
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
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
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
                    Exit Do
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                Exit Function
        End Select
        
        
        If StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
            StrConv(ZAIKOREC.Retu, vbUnicode) & _
            StrConv(ZAIKOREC.Ren, vbUnicode) & _
            StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(GOODSREC.ST_SOKO, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_RETU, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_REN, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_DAN, vbUnicode) Then
        
        
            For i = 0 To UBound(EE_ZAIKO_TBL)
                            
                If Trim(EE_ZAIKO_TBL(i).EE_LOC) = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
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
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function

Public Function F106021_Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional LOCATION As String = "        ") As Integer
'****************************************************
'*      在庫数集計
'*
'*  品番または品番＋棚番毎の在庫数を集計する。
'*
'*  引数 :  在庫数（商品化済み）
'*          在庫数（未商品）
'*          事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番(省略可 省略=空白)
'*  戻り値: false    正常
'*          SYS_ERR  継続できない異常
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Soko_No     As String * 2
Dim Retu        As String * 2
Dim Ren         As String * 2
Dim Dan         As String * 2
    
Dim Not_GOODS   As Boolean

Dim i           As Integer

    F106021_Zaiko_Syukei_Proc = SYS_ERR

    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreater

    If Len(Trim(LOCATION)) = 0 Then
                                '倉庫番号空白は棚番省略とみなす
        Call UniCode_Conv(K1_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K1_ZAIKO.Retu, "")
        Call UniCode_Conv(K1_ZAIKO.Ren, "")
        Call UniCode_Conv(K1_ZAIKO.Dan, "")

        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop

    Else

        Soko_No = Mid(LOCATION, 1, 2)
        Retu = Mid(LOCATION, 3, 2)
        Ren = Mid(LOCATION, 5, 2)
        Dan = Mid(LOCATION, 7, 2)

        Call UniCode_Conv(K0_ZAIKO.Soko_No, Soko_No)
        Call UniCode_Conv(K0_ZAIKO.Retu, Retu)
        Call UniCode_Conv(K0_ZAIKO.Ren, Ren)
        Call UniCode_Conv(K0_ZAIKO.Dan, Dan)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(Retu)) = 0 Then
                        Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
                    End If
                    If Len(Trim(Ren)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
                    End If
                    If Len(Trim(Dan)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Dan, vbUnicode)
                    End If

                    If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫データ")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop
    End If

    F106021_Zaiko_Syukei_Proc = False

End Function



Private Function Data_Make_Sub() As Integer
    
Dim sts         As Integer
Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
Dim AVE_QTY     As Long
Dim ans         As Integer
    
    
    Data_Make_Sub = True
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))


    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Data_Make_Sub = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select

    '-----------------------------------------  '商品化集計ファイル作成
                                                '在庫集計処理
    If F106021_Zaiko_Syukei_Proc(Sumi_QTY, _
                            Mi_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
        Exit Function
    End If
                                                
                                                
    If Sumi_QTY > 0 Then
        Sumi_QTY = Sumi_QTY - 1             'サンプル分マイナス
    End If
                                            '商品化済み在庫数
    Call UniCode_Conv(GOODSREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                            '未商品在庫数
    Call UniCode_Conv(GOODSREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                    
                                            '月平均出荷数
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    
    AVE_QTY = 0
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    Select Case sts
        Case BtNoErr
            Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
            AVE_QTY = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
        Case BtErrKeyNotFound
            Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
            Exit Function
    End Select
    Call UniCode_Conv(GOODSREC.AVE_SYUKA, Format(AVE_QTY, "00000000"))
                                            '事前商品化状況
    If AVE_QTY = 0 Then
        Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")
    Else
        Call UniCode_Conv(GOODSREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
    End If
        
        
    Do
        
        sts = BTRV(BtOpUpdate, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
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

    Data_Make_Sub = False


End Function

Private Function Ukeharai_Set_Proc() As Integer

Dim com As Integer
Dim sts As Integer

    
    
    Ukeharai_Set_Proc = True
    
    
    
    Combo(pcmbUKEHARAI_CODE).Clear
    Combo(pcmbUKEHARAI_CODE).AddItem "全　て　　　　　　　　　　　　" & "     " & " "
    
    
    
    
    Call UniCode_Conv(K1_P_UKEHARAI.TORI_KBN, Right(Combo(pcmbTORI_KBN), 1))
    Call UniCode_Conv(K1_P_UKEHARAI.UKEHARAI_CODE, "")


    com = BtOpGetGreaterEqual


    Do
        DoEvents
        
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K1_P_UKEHARAI, Len(K1_P_UKEHARAI), 1)
        Select Case sts
            Case BtNoErr
                If Trim(Right(Combo(pcmbTORI_KBN), 1)) <> "" Then
                    If Right(Combo(pcmbTORI_KBN), 1) <> StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) Then
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        End Select
        
        
        Combo(pcmbUKEHARAI_CODE).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & _
                                            "     " & StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    Loop

    Combo(pcmbUKEHARAI_CODE).ListIndex = 0


    Ukeharai_Set_Proc = False


End Function
