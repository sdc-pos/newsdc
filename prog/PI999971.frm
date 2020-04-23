VERSION 5.00
Begin VB.Form PI999971 
   Caption         =   "商品ラベル一括発行 ([PI99997]　2011.08.05 11:00"
   ClientHeight    =   6264
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10188
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
   ScaleHeight     =   6264
   ScaleWidth      =   10188
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text2 
      Height          =   3855
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   5412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｸﾘｱｰ"
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
      Left            =   420
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   9
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   8
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   0
      Left            =   4620
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   3888
      Left            =   6156
      TabIndex        =   2
      Top             =   1200
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
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
      Left            =   8772
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷"
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
      Left            =   7416
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ＮＧ件数"
      Height          =   252
      Index           =   4
      Left            =   7932
      TabIndex        =   7
      Top             =   5760
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "ＯＫ件数"
      Height          =   252
      Index           =   3
      Left            =   7932
      TabIndex        =   6
      Top             =   5280
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "読込み件数"
      Height          =   252
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   5280
      Width           =   1272
   End
   Begin VB.Label Label1 
      Caption         =   "印刷結果"
      Height          =   252
      Index           =   1
      Left            =   6156
      TabIndex        =   3
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "PI999971"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI999971.MousePointer = vbHourglass

    Call Ctrl_Lock(PI999971)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI999971)


    PI999971.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim yn          As Integer
    
    
    Error_Check_Proc = True
    
        
        
    Error_Check_Proc = False
    

End Function

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer




Dim com             As Integer
Dim sts             As Integer

Dim Skip_F          As Boolean

Dim wkLine          As Variant
Dim wkItem          As Variant


Dim Parts_F             As Integer
Dim Gaisou_F            As Integer
Dim Kishu_F             As Integer
Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long



'=============================== 2011.08.04 =====
Dim Parts       As String   '品番
Dim ID          As Long     '指示№

Dim PartsLabel  As Integer  '品番ﾗﾍﾞﾙ 0:なし 以外：枚数
Dim KisyuLabel  As Integer  '機種ﾗﾍﾞﾙ 0:なし
Dim JanLabel    As Integer  'JANﾗﾍﾞﾙ 0:なし
Dim GLabel      As Integer  '外装ﾗﾍﾞﾙ 0:なし
Dim ItemLabel   As Integer  'ｱｲﾃﾑﾗﾍﾞﾙ枚数

Dim OrderNo     As String
Dim ItemNo      As String

Dim Pri_Date    As String

Dim L_QTY       As Long

Dim wkDate      As String * 10
Dim NGItem      As String * 20
'=============================== 2011.08.04 =====




Dim objAccess       As Access.Application
Dim strAccessPath   As String



    Select Case Index
        Case 0              '印刷
            
            
            
            Beep
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                
                List2.Clear
                
                OK_cnt = 0
                NG_cnt = 0
                Text1(1).Text = Format(OK_cnt, "#,##0")
                Text1(2).Text = Format(NG_cnt, "#,##0")
                
                    
                    
                                    
                wkLine = Split(Text2.Text, vbCrLf, -1)
                                    
                                    
                For i = 0 To UBound(wkLine)
                    
                    
                    wkItem = Split(wkLine(i), vbTab, -1)
                    
                    Skip_F = False
                    
                    If UBound(wkItem) < 2 Then
                    Else
                        
                        
                        If Not IsNumeric(wkItem(2)) Then
                            Skip_F = True
                        Else
                            If CInt(wkItem(2)) = 0 Then
                                Skip_F = True
                            End If
                        End If
                        
                        wkDate = ""
                        If UBound(wkItem) >= 3 Then
                            If Trim(wkItem(3)) <> "" Then
                                If Not IsDate(wkItem(3)) Then
                                    Skip_F = True
                                Else
                                    wkDate = Format(wkItem(3), "YYYY/MM/DD")
                                End If
                            End If
                        End If
                                                                        
                                                                        
                        If Not Skip_F Then
                        
                        
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Format(wkItem(0)))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Format(wkItem(1)))
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    
                                    Skip_F = True
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Sub
                            
                            End Select
                        End If
                
                
                
                
                        If Skip_F Then
                            
                            NGItem = wkItem(1)
                            List2.AddItem wkItem(0) & " " & NGItem & " " & "NG"
                            
                            NG_cnt = NG_cnt + 1
                            Text1(2).Text = Format(NG_cnt, "#,##0")
                        Else
                        
                        
                                On Error Resume Next
                                Set objAccess = GetObject(, "Access.Application")
                                If Err().Number <> 0 Then
                                    MsgBox "この端末では商品ラベル発行は行えません。"
            '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                                Else
            '                        MsgBox Err.Number
                                    
                                    strAccessPath = App.Path
                                    If Right(strAccessPath, 1) <> "\" Then
                                        strAccessPath = strAccessPath & "\"
                                    End If
                                    
                                    strAccessPath = strAccessPath & "litem.mdb"
                                    Set objAccess = GetObject(strAccessPath)
                
                                
                        
                                    Parts_F = 1
                                    
                                    
                                    Gaisou_F = 0
                                    
                                    Kishu_F = 0
                                    
                                    GAISOU_QTY = 0
                                    
                                    GAISOU_SHIJI_QYU = 0
                                    
                                    com = BtOpGetFirst
                                    Do
                                    
                                    
                                    
                                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                                
                                                Select Case sts
                                                
                                                    Case BtNoErr
                                                    Case Else
                                                        Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                                        Exit Sub
                                                End Select
                                            
                                            Case BtErrEOF
                                                Exit Do
                                            Case Else
                                                Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                                Exit Sub
                                        End Select
                                        
                                        com = BtOpGetNext
                                    
                                    
                                    Loop
                                    
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Format(wkItem(0)))
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Format(wkItem(1)))
                            
                            
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            
                                            Call UniCode_Conv(ITEMREC.L_IRI_QTY, "00000000")
                                            
                                            
                                            sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                            Select Case sts
                                                Case BtNoErr
                                        
                                                    
                                                    
                                                    '=================================================2011.08.04
'                                                    objAccess.Run "PosPrintLabel", Trim(Format(wkItem(1))), CLng(wkItem(2)), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0
                                        
                                        
                                        
                                                    PartsLabel = 0
                                                    KisyuLabel = 0
                                                    JanLabel = 0
                                                    GLabel = 0
                                                    ItemLabel = 0


                                                    '品目コード
                                                    Parts = wkItem(1)
                                                    'パーツラベル
                                                    PartsLabel = CLng(wkItem(2))
                                                    'ID
                                                    ID = 0
                                                    'アイテムラベル
                                                    ItemLabel = 0
                                                    'オーダー№
                                                    OrderNo = ""
                                                    'アイテム№
                                                    ItemNo = ""
                                                    '印刷日付
                                                    Pri_Date = wkDate
                                                    '数量
                                                    L_QTY = 1
                                                    objAccess.Run "NewPosPrintLabel", _
                                                                        Trim(Parts), _
                                                                        PartsLabel, _
                                                                        KisyuLabel, _
                                                                        JanLabel, _
                                                                        GLabel, _
                                                                        ID, _
                                                                        ItemLabel, _
                                                                        Trim(OrderNo), _
                                                                        Trim(ItemNo), _
                                                                        Pri_Date, _
                                                                        L_QTY
                                                    '=================================================2011.08.04
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                                Case Else
                                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                                    Exit Sub
                                                
                                        
                                            End Select
                                        
                                        Case BtErrKeyNotFound
                                            
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Sub
                                    
                                    End Select
                                    
                                    
                                    
                                    
                                                            
                                    Set objAccess = Nothing
                                End If
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                            OK_cnt = OK_cnt + 1
                            Text1(1).Text = Format(OK_cnt, "#,##0")
                        
                        
                        End If
            
                    End If
                Next i
            
                MsgBox "印刷が終了しました。"
            
            
            End If
        Case 1              '終了
            Unload Me
    
        Case 2
            Text2.Text = ""
    End Select

ErrHandler:
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant


    If App.PrevInstance Then
        MsgBox "同一プログラム実行中です。"
        End
    End If

                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                'ラベル用品目マスタＯＰＥＮ
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    
    
    
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
    
                                            'ラベル用品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ラベル用品目マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI999971 = Nothing

    End
End Sub


Private Sub RichTextBox1_Change()

End Sub

