VERSION 5.00
Begin VB.Form F1011501 
   BackColor       =   &H00FFFFFF&
   Caption         =   "振替品番 メンテナンス"
   ClientHeight    =   7500
   ClientLeft      =   2130
   ClientTop       =   2730
   ClientWidth     =   14295
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
   ScaleHeight     =   7500
   ScaleWidth      =   14295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   12960
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2400
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   11520
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   8040
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2400
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   9840
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2400
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      Index           =   1
      Left            =   1740
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1320
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      Index           =   3
      Left            =   1740
      TabIndex        =   4
      Top             =   2400
      Width           =   4920
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   360
      Index           =   2
      Left            =   1740
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1860
      Width           =   2520
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1740
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   345
      Index           =   0
      Left            =   1740
      MaxLength       =   5
      OLEDragMode     =   1  '自動
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.ListBox List1 
      Height          =   2940
      ItemData        =   "F1011501.frx":0000
      Left            =   840
      List            =   "F1011501.frx":0002
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   12915
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
      Left            =   10260
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   9420
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   8580
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   7740
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   6420
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   5580
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   4740
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6660
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "削 除"
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
      Left            =   3900
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6660
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "表 示"
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
      Left            =   2580
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   1740
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6660
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
      Left            =   900
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6660
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更 新"
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
      Left            =   60
      TabIndex        =   9
      Top             =   6660
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "員数"
      Height          =   240
      Index           =   7
      Left            =   12360
      TabIndex        =   34
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "有効長"
      Height          =   240
      Index           =   6
      Left            =   10680
      TabIndex        =   33
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "切断前長さ"
      Height          =   240
      Index           =   5
      Left            =   6720
      TabIndex        =   32
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "切断数"
      Height          =   240
      Index           =   4
      Left            =   9000
      TabIndex        =   31
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   30
      Top             =   1860
      Width           =   5025
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   29
      Top             =   1320
      Width           =   5025
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "備　考"
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   28
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "振替先品番"
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   27
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "振替元品番"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   26
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label LabJIGYO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   240
      Left            =   120
      TabIndex        =   25
      Top             =   6300
      Width           =   2475
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   720
      TabIndex        =   24
      Top             =   900
      Width           =   780
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   23
      Top             =   120
      Width           =   5025
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "担当者"
      Height          =   255
      Index           =   1
      Left            =   660
      TabIndex        =   22
      Top             =   180
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
Attribute VB_Name = "F1011501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    '設定用ファイル （CSV)
Dim In_Files    As String   'C:\SDC\FILES\FURIKAE.CSV

Dim W_Edit      As String

Dim W_Disp_Key  As String

Private Const ptxTANTO% = 0
Private Const ptxMOTO% = 1
Private Const ptxSAKI% = 2
Private Const ptxBIKOU% = 3

Private Const ptxCUT_SU% = 4        '2012.03.14

Private Const ptxMOTO_LEN% = 5      '2012.12.26

Private Const ptxYUKO_LEN% = 6      '2012.12.26

Private Const ptxKO_QTY% = 7        '2013.02.22


Private Const Text_Max% = 7         '2013.02.22 6-->7

Private Const plbTANTO% = 0
Private Const plbMOTO% = 1
Private Const plbSAKI% = 2

Private Const fncDO% = 0
Private Const fncDISP% = 3
Private Const fncDEL% = 4
Private Const fncEND% = 11



Private Const pcmbNAIGAI% = 0       '2012.03.13

'Private Const pcmbNAIGAI% = 0

Private Const LAST_UPDATE_DAY$ = "(F101150 2012.03.24 16:15)"

Private Function DEL_CHK()

Dim sts         As Integer
Dim yn          As Integer
    
    DEL_CHK = True
    
    If Trim(Text1(ptxMOTO)) = "" Then
        MsgBox "振替元が未指定エラー！"
        Text1(ptxMOTO).SetFocus
        Call Text1_GotFocus(ptxMOTO)
        Exit Function
    End If
        
    If Trim(Text1(ptxSAKI)) = "" Then
        MsgBox "振替先が未指定エラー！"
        Text1(ptxSAKI).SetFocus
        Call Text1_GotFocus(ptxSAKI)
        Exit Function
    End If
        
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                           '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))     '2012.03.13
    
    
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, Trim(Text1(ptxMOTO)))
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_GO, Last_JGYOBU)                           '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_GO, Right(Combo1(pcmbNAIGAI).Text, 1))     '2012.03.13
    
    
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, Trim(Text1(ptxSAKI)))
    Do
        sts = BTRV(BtOpGetEqual, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                
                MsgBox "指定されたマスターがありません。"
                Text1(ptxMOTO).SetFocus
                Call Text1_GotFocus(ptxMOTO)
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                            "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
            
    
    
    DEL_CHK = False

End Function

Private Function MST_DEL()
Dim sts         As Integer
Dim yn          As Integer
    
    MST_DEL = True
        
    
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                           '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, Trim(Text1(ptxMOTO)))
    
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_GO, Last_JGYOBU)                           '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_GO, Right(Combo1(pcmbNAIGAI).Text, 1))     '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, Trim(Text1(ptxSAKI)))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                
                MsgBox "指定されたマスターがありません。"
                Text1(ptxMOTO).SetFocus
                Call Text1_GotFocus(ptxMOTO)
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                yn = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                            "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
    
    Do
        sts = BTRV(BtOpDelete, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpDelete, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
    
    
    
    MST_DEL = False
    
End Function
Private Sub List_Proc()
'----------------------------------------------------------------------------
'                   リストボックス表示処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer
    
    
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                           '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))     '2012.03.13
    
    
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, W_Disp_Key)
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, "")
       
    com = BtOpGetGreaterEqual
    Do
        
        DoEvents
        
        Do
            sts = BTRV(com, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
                
            Select Case sts
                Case BtNoErr
                    
                    '---------------------------------------------  2012.03.13
                    If StrConv(FURIKAEREC.JGYOBU_MAE, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(FURIKAEREC.NAIGAI_MAE, vbUnicode) <> Right(Combo1(pcmbNAIGAI).Text, 1) Then
                        
                        sts = BtErrEOF
                        
                        Exit Do
                    End If
                    '---------------------------------------------  2012.03.13
                                    
                    Exit Do
                
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Call FURIKAE_CLR
                        'MsgBox "指定された工程がありません。"
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    ans = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                                "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If ans = vbNo Then Exit Do
                Case Else
                    Call File_Error(sts, com, "品番振替Ｍ")
                    Exit Do
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        If Trim(Text1(ptxMOTO)) <> "" Then
            If Left(StrConv(FURIKAEREC.HIN_MAE, vbUnicode), Len(W_Disp_Key)) <> W_Disp_Key Then
                Exit Do
            End If
        End If
        
        Call List_Edit
        List1.AddItem W_Edit
    
        com = BtOpGetNext
    Loop
    

End Sub

Private Sub List_Edit()
    
    W_Edit = ""
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.HIN_MAE, vbUnicode)
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.HIN_GO, vbUnicode)
    
    W_Edit = W_Edit & StrConv(FURIKAEREC.BIKOU, vbUnicode)

End Sub
Private Sub Clear_Field(Start As Integer)
'----------------------------------------------------------------------------
'                   画面消去処理
'----------------------------------------------------------------------------
Dim i As Integer

    For i = Start To Text_Max%
        Text1(i).Text = ""
    Next i
    Label1(0).Caption = ""

End Sub
Private Function Error_Check_Proc(Index As Integer, Chk_Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim ans     As Integer
    
    Error_Check_Proc = True
    
    Select Case Index
    
        Case ptxTANTO%    '担当者
            
            If Trim(Text1(Index)) = "" Then
                MsgBox "担当者　未設定エラー！"
                Exit Function
            End If
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO%).Text)
    
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            
            Select Case sts
                Case BtNoErr
                    Label1(0).Caption = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                Case BtErrKeyNotFound
                
''                    If Chk_Mode = 0 Then
''                        Label1(0).Caption = "担当者コード　未登録"
''                    Else
''                        MsgBox "入力した項目はエラーです（担当者 未登録）"
''                        Text1(Index).SetFocus
''                        Exit Function
''                    End If
                    '2012.12.25     上記を下記に変更        M.T
                    Select Case Chk_Mode
                        Case 0
                            Label1(0).Caption = "担当者コード　未登録"
                        Case 1
                            MsgBox "入力した項目はエラーです（担当者 未登録）"
                            Text1(Index).SetFocus
                            Exit Function
                        Case Else
                            MsgBox "担当者コードが未登録の為、更新や削除は出来ません。"
                            
                            Exit Function
                    End Select
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>     ここまで。
                    
                    
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
        
        Case ptxMOTO, ptxSAKI
            If Trim(Text1(Index)) = "" Then
                If Index = ptxMOTO Then
                    MsgBox "振替元品番　未設定エラー！"
                Else
                    MsgBox "振替先品番　未設定エラー！"
                End If
                Exit Function
            End If
            
            
            '2012.12.25 追加    M.T
            'If Index = ptxSAKI Then
                Text1(Index) = UCase(Text1(Index))
                DoEvents
            'End If
            
            
            
            
                
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
 '2012.03.13           Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)              '国内
            Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Index))
            Do
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Label1(Index) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        Exit Do
                    Case BtErrKeyNotFound       'レコード無し
                        If Chk_Mode = 0 Then
                            Label1(Index).Caption = "品番　未登録"
                            
                            
                        Else
                            MsgBox "入力した項目はエラーです（品番 未登録）"
                            Text1(Index).SetFocus
                            
                            Exit Function
                        End If
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        ans = MsgBox("他で使用中です！<品番Ｍ>" & Chr(13) & Chr(10) & _
                                    "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If ans = vbNo Then Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品番Ｍ")
                        Exit Function
                End Select
            Loop
            
            If Index = ptxSAKI Then
                Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                       '2012.03.13
                Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
                Call UniCode_Conv(K0_FURIKAE.HIN_MAE, Text1(ptxMOTO))
                
                Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                       '2012.03.13
                Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
                Call UniCode_Conv(K0_FURIKAE.HIN_GO, Text1(ptxSAKI))
                Do
                    sts = BTRV(BtOpGetEqual, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
                    
                    Select Case sts
                        Case BtNoErr
                            If Trim(Text1(ptxBIKOU)) = "" Then
                                Text1(ptxBIKOU) = Trim(StrConv(FURIKAEREC.BIKOU, vbUnicode))
                            End If
                            
                            '2012.03.14
                            If Trim(Text1(ptxCUT_SU).Text) = "" Then
                                If IsNumeric(StrConv(FURIKAEREC.CUT_SU, vbUnicode)) Then
                                    Text1(ptxCUT_SU).Text = Format(Val(StrConv(FURIKAEREC.CUT_SU, vbUnicode)), "#")
                                End If
                            End If
                            '2012.03.14
                            
                            '2012.12.26
                            If Trim(Text1(ptxMOTO_LEN).Text) = "" Then
                                If IsNumeric(StrConv(FURIKAEREC.MOTO_LEN, vbUnicode)) Then
                                    Text1(ptxMOTO_LEN).Text = Format(Val(StrConv(FURIKAEREC.MOTO_LEN, vbUnicode)), "#")
                                End If
                            End If
                            Text1(ptxYUKO_LEN).Text = ToRoundDown(CCur(Val(Text1(ptxMOTO_LEN).Text) / Val(Text1(ptxCUT_SU).Text)), 0) '2013.01.25
                            '2012.12.26
                            
                            
                            
                            Exit Do
                        Case BtErrKeyNotFound       'レコード無し
                            
                            'MsgBox "指定された工程がありません。"
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            ans = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                                        "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                            If ans = vbNo Then Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品番振替Ｍ")
                            Exit Function
                    End Select
                Loop
            
            
            
            End If
        
        Case ptxBIKOU
            Call UniCode_Conv(FURIKAEREC.BIKOU, Text1(ptxBIKOU))
            
            If Trim(Text1(ptxBIKOU)) <> Trim(StrConv(FURIKAEREC.BIKOU, vbUnicode)) Then
                MsgBox "文字数エラー!"
                Exit Function
            End If
            
            
        Case ptxCUT_SU                      '2012.03.14
            
            If Last_JGYOBU = SHIZAI Or Last_JGYOBU = BUZAI Then
                If Not IsNumeric(Text1(ptxCUT_SU).Text) Then
                    MsgBox "入力した項目はエラーです（切断数）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            
                If Val(Text1(ptxCUT_SU).Text) < 1 Then
                    MsgBox "入力した項目はエラーです（切断数）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            
                Text1(ptxYUKO_LEN).Text = ToRoundDown(CCur(Val(Text1(ptxMOTO_LEN).Text) / Val(Text1(ptxCUT_SU).Text)), 0) '2013.01.25
            
            
            Else
                Text1(ptxCUT_SU).Text = ""
            End If
        
        
        Case ptxMOTO_LEN                    '2012.12.26
            
            If Last_JGYOBU = SHIZAI Or Last_JGYOBU = BUZAI Then
                If Not IsNumeric(Text1(ptxMOTO_LEN).Text) Then
                    MsgBox "入力した項目はエラーです（切断前長さ）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            
                If Val(Text1(ptxMOTO_LEN).Text) < 1 Then
                    MsgBox "入力した項目はエラーです（切断前長さ）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            
                If Val(Text1(ptxCUT_SU).Text) = 0 Then
                Else
                    Text1(ptxYUKO_LEN).Text = ToRoundDown(CCur(Val(Text1(ptxMOTO_LEN).Text) / Val(Text1(ptxCUT_SU).Text)), 0)
                End If
            Else
                Text1(ptxMOTO_LEN).Text = ""
            End If
        
        Case ptxKO_QTY                      '2013.02.22
        
            If Last_JGYOBU = SHIZAI Or Last_JGYOBU = BUZAI Then
                If Not IsNumeric(Text1(ptxKO_QTY).Text) Then
                    MsgBox "入力した項目はエラーです（員数）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            
                            
                Text1(ptxKO_QTY).Text = Format(Val(Text1(ptxKO_QTY).Text), "0.00")
            
            
                If Val(Text1(ptxKO_QTY).Text) < 0 Then
                    MsgBox "入力した項目はエラーです（員数）"
                    Text1(Index).SetFocus
                    Exit Function
                End If
            Else
                Text1(ptxKO_QTY).Text = ""
            End If
                
        
        
        Case Else
    
        
    End Select
        
    Error_Check_Proc = False
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   追加／変更処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim W_MSG           As String


Dim W_Date          As String
Dim W_Time          As String

    Update_Proc = True
    
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
     "更新処理開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_MAE, Last_JGYOBU)                       '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_MAE, Trim(Text1(ptxMOTO)))
    
    Call UniCode_Conv(K0_FURIKAE.JGYOBU_GO, Last_JGYOBU)                       '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.NAIGAI_GO, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
    Call UniCode_Conv(K0_FURIKAE.HIN_GO, Trim(Text1(ptxSAKI)))
    Do
        sts = BTRV(BtOpGetEqual, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
        
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                                
                                
                Exit Do
            Case BtErrKeyNotFound       'レコード無し
                Call FURIKAE_CLR
                
                Call UniCode_Conv(FURIKAEREC.JGYOBU_MAE, Last_JGYOBU)                       '2012.03.13
                Call UniCode_Conv(FURIKAEREC.NAIGAI_MAE, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
                Call UniCode_Conv(FURIKAEREC.HIN_MAE, Trim(Text1(ptxMOTO)))
                
                Call UniCode_Conv(FURIKAEREC.JGYOBU_GO, Last_JGYOBU)                       '2012.03.13
                Call UniCode_Conv(FURIKAEREC.NAIGAI_GO, Right(Combo1(pcmbNAIGAI).Text, 1))  '2012.03.13
                Call UniCode_Conv(FURIKAEREC.HIN_GO, Trim(Text1(ptxSAKI)))
                
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                ans = MsgBox("他で使用中です！<FURIKAE>" & Chr(13) & Chr(10) & _
                            "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                If ans = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
    
    Call UniCode_Conv(FURIKAEREC.BIKOU, Trim(Text1(ptxBIKOU)))
    
    '2012.03.14
    If IsNumeric(Text1(ptxCUT_SU).Text) Then
        Call UniCode_Conv(FURIKAEREC.CUT_SU, Format(Val(Text1(ptxCUT_SU).Text), "000"))
    Else
        Call UniCode_Conv(FURIKAEREC.CUT_SU, "")
    End If
    '2012.03.14
    
    
    '2012.12.26
    If IsNumeric(Text1(ptxMOTO_LEN).Text) Then
        Call UniCode_Conv(FURIKAEREC.MOTO_LEN, Format(Val(Text1(ptxMOTO_LEN).Text), "000"))
    Else
        Call UniCode_Conv(FURIKAEREC.MOTO_LEN, "")
    End If
    '2012.12.26
    
    
    
    '2012.02.25
    If IsNumeric(Text1(ptxKO_QTY).Text) Then
        Call UniCode_Conv(FURIKAEREC.KO_QTY, Format(Val(Text1(ptxKO_QTY).Text), "0.00"))
    Else
        Call UniCode_Conv(FURIKAEREC.KO_QTY, "")
    End If
    '2012.02.25
    
    
    
    
    W_Date = Format(Date, "yyyymmdd")
    W_Time = Format(Time, "hhmmss")
    If com = BtOpUpdate Then
        Call UniCode_Conv(FURIKAEREC.UPD_TANTO, Trim(Text1(ptxTANTO)))
        Call UniCode_Conv(FURIKAEREC.UPD_DATETIME, W_Date & W_Time)
        
    Else
        Call UniCode_Conv(FURIKAEREC.INS_TANTO, Trim(Text1(ptxTANTO)))
        Call UniCode_Conv(FURIKAEREC.Ins_DateTime, W_Date & W_Time)
        
    End If
    
    Do
        sts = BTRV(com, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
        
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                Sleep (500)
            Case Else
                Call File_Error(sts, com, "品番振替Ｍ")
                Exit Function
        End Select
    Loop
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
     "更新処理終了" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
     
     
    Update_Proc = False
    
    
End Function


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
Dim X_i     As Integer

    Select Case Index
        Case fncDO
            For X_i = ptxTANTO To ptxCUT_SU
                If Error_Check_Proc(X_i, 0) Then    'エラーチェック
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
                      
            Next X_i
                                       
                                       
            '2012.12.25     追加        M.T
            If Error_Check_Proc(ptxTANTO, 2) Then     'エラーチェック
                Text1(ptxTANTO).SetFocus
                Call Text1_GotFocus(ptxTANTO)
                Exit Sub
            End If
            '>>>>>>>>>>>>>  ここまで。
                                            
                                            
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbNo Then
                Command(fncEND).SetFocus
                Exit Sub
            End If
            
            
            
            Call Input_Lock
            If Update_Proc() Then
                Unload Me
            End If
            Call Input_UnLock
            MsgBox "更新しました。"
                
            Call Clear_Field(ptxSAKI)
            List1.Clear
            Call List_Proc
            
            
            Text1(ptxSAKI) = ""
            Label1(plbSAKI) = ""
            Text1(ptxSAKI).SetFocus
            Call Text1_GotFocus(ptxSAKI)
            
            'Command(fncEND).SetFocus
            
        Case fncDISP
            If Error_Check_Proc(X_i, 0) Then    'エラーチェック
                Text1(ptxTANTO).SetFocus
                Call Text1_GotFocus(ptxTANTO)
                Exit Sub
            End If
            
            W_Disp_Key = Trim(Text1(ptxMOTO))
            List1.Clear
            Call List_Proc
            
            Text1(ptxMOTO) = W_Disp_Key
            Label1(plbMOTO) = ""
            Text1(ptxSAKI) = ""
            Label1(plbSAKI) = ""
            
            Text1(ptxBIKOU) = ""
                
            If List1.ListCount < 1 Then
                
                Text1(ptxMOTO).SetFocus
            Else
                
                List1.SetFocus
            End If
            
        Case fncDEL
            If DEL_CHK Then
                'Text1(ptxMOTO).SetFocus
                'Call Text1_GotFocus(ptxMOTO)
                Exit Sub
            End If
            
            '2012.12.25     追加        M.T
            If Error_Check_Proc(ptxTANTO, 2) Then     'エラーチェック
                Text1(ptxTANTO).SetFocus
                Call Text1_GotFocus(ptxTANTO)
                Exit Sub
            End If
            '>>>>>>>>>>>>>  ここまで。
            
            
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbNo Then
                Text1(ptxMOTO).SetFocus
                Call Text1_GotFocus(ptxMOTO)
                Exit Sub
            End If
            
            If MST_DEL Then
                MsgBox "マスター削除エラー！"
                Unload Me
            End If
            MsgBox "削除しました。"
            List1.Clear
            Call List_Proc
            Call Clear_Field(ptxSAKI)
            
            
            Text1(ptxBIKOU) = ""
            Label1(plbSAKI) = ""
            Text1(ptxSAKI) = ""
            Text1(ptxSAKI).SetFocus
            
            
        Case fncEND
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub


Private Sub Form_DblClick()
'    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
        Case vbKeyZ
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

    Show
    
    Call Clear_Field(0)
    DoEvents
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                
                                
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "振替品番ﾏｽﾀﾒﾝﾃﾅﾝｽ" & LAST_UPDATE_DAY, Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                
                                
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
            F1011501.Caption = "振替品番ﾏｽﾀﾒﾝﾃﾅﾝｽ（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
    Combo1(pcmbNAIGAI).AddItem NAIGAI1$ & "   " & NAIGAI_NAI$
    Combo1(pcmbNAIGAI).AddItem NAIGAI2$ & "   " & NAIGAI_GAI$
    Combo1(pcmbNAIGAI).ListIndex = 0
                                
                                
                                
                                    
                                '振替品番マスタＯＰＥＮ
    If FURIKAE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
'    If GetIni("F101150", "IN_FILE", "F101150", c) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        End
'    End If
'    In_Files = RTrim(c)
'
'    Text1(ptxFILE%) = In_Files
'
'    Command(fncDO).Enabled = False
    c = ""
    'Call List_Proc
    W_Disp_Key = ""
    Text1(ptxTANTO).SetFocus
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    Call FURIKAE_CLOSE
        
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品番Ｍ")
        End If
    End If
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者Ｍ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, FURIKAE_POS, FURIKAEREC, Len(FURIKAEREC), K0_FURIKAE, Len(K0_FURIKAE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1011501 = Nothing
    End
    
End Sub





Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1011501.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011501)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1011501)

    F1011501.MousePointer = vbDefault

End Sub

Private Sub List1_DblClick()
Dim W_ITEM          As String
Dim X_i             As Integer


    If List1.ListIndex < 0 Then List1.ListIndex = 0
    W_ITEM = Left(List1.List(List1.ListIndex), UBound(FURIKAEREC.HIN_MAE) + 1)
    
    Text1(ptxMOTO) = W_ITEM
    
    W_ITEM = Mid(List1.List(List1.ListIndex), UBound(FURIKAEREC.HIN_MAE) + 1 + 1, UBound(FURIKAEREC.HIN_GO) + 1)
    
    Text1(ptxSAKI) = W_ITEM
    
    If Error_Check_Proc(ptxMOTO, 0) Then    'エラーチェック
        Text1(ptxMOTO).SetFocus
        Call Text1_GotFocus(ptxMOTO)
        Exit Sub
    End If
    Text1(ptxBIKOU) = ""
    
    Text1(ptxCUT_SU) = ""                   '2012.05.19
    
    
    Text1(ptxMOTO_LEN) = ""                 '2012.12.26
    
    
    If Error_Check_Proc(ptxSAKI, 0) Then    'エラーチェック
        Text1(ptxSAKI).SetFocus
        Call Text1_GotFocus(ptxSAKI)
        Exit Sub
    End If
    
    Text1(ptxBIKOU).SetFocus
    Call Text1_GotFocus(ptxBIKOU)
    
'    If ZAIKO_Get(W_ITEM, 0) Then Exit Sub
'    KG50102.Show vbModal
    
End Sub

Private Sub List1_GotFocus()
    If List1.ListCount > 0 And _
       List1.ListIndex < 0 Then
        List1.ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
    If List1.ListCount <= 0 Then Exit Sub
    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)     '移動
    Else
        Call List1_DblClick
    End If


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
    F1011501.Caption = "振替品番ﾏｽﾀﾒﾝﾃﾅﾝｽ（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index, 0) Then    'エラーチェック      'エラー：警告
    
    'If Error_Check_Proc(Index, 1) Then    'エラーチェック       'エラー：再入力
        
        Text1(Index).SetFocus
        Call Text1_GotFocus(Index)
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Text1_LostFocus(Index As Integer)
    '2012.12.25     追加    M.T
    Text1(ptxMOTO) = UCase(Text1(ptxMOTO))
    Text1(ptxSAKI) = UCase(Text1(ptxSAKI))
    DoEvents
    
End Sub
' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
'
'
'       2012.03.25  frm より　移管
'
' ------------------------------------------------------------------------
Private Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
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

