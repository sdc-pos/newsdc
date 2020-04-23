VERSION 5.00
Begin VB.Form Conv_ZAIKO1 
   Caption         =   "在庫仕入先／単価処理"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12645
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
   ScaleHeight     =   6300
   ScaleWidth      =   12645
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   10680
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   9000
      MaxLength       =   11
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   8160
      MaxLength       =   8
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Index           =   0
      ItemData        =   "Conv_ZAIKO1.frx":0000
      Left            =   240
      List            =   "Conv_ZAIKO1.frx":0002
      TabIndex        =   11
      Top             =   1800
      Width           =   10455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   23
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   2640
      TabIndex        =   15
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "仕入先"
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "数量"
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   32
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "入荷日"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   31
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "段"
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "連"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "列"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "仕入単価"
      Height          =   255
      Index           =   4
      Left            =   9600
      TabIndex        =   27
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "倉庫"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   26
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblBikou 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   4560
      TabIndex        =   25
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "品番"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Conv_ZAIKO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxHIN_GAI% = 0               '品番
Private Const ptxSoko_No% = 1               '倉庫
Private Const ptxRetu% = 2                  '列
Private Const ptxRen% = 3                   '連
Private Const ptxDan% = 4                   '段
Private Const ptxNYUKA_YY% = 5              '入荷日（年）
Private Const ptxNYUKA_MM% = 6              '入荷日（月）
Private Const ptxNYUKA_DD% = 7              '入荷日（日）
Private Const ptxYUKO_Z_QTY% = 8            '数量
Private Const ptxSHIIRE_CODE% = 9           '仕入先
Private Const ptxSHIIRE_TANKA% = 10         '仕入単価

Private Const ptxGOODS_ON% = 11             '商品／未商品

'リストＢＯＸ用添字
Private Const plstNYU% = 0



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    Conv_ZAIKO1.MousePointer = vbHourglass

    Call Ctrl_Lock(Conv_ZAIKO1)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(Conv_ZAIKO1)


    Conv_ZAIKO1.MousePointer = vbDefault

End Sub

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   リストボックス表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim wkQty       As String
Dim wkTanka     As String

Dim List        As String


    List_Disp_Proc = True
    
    List1(plstNYU).Clear
    
    
    
    '資材分在庫データ読み込み
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")

    
    com = BtOpGetGreater
    
    
    Do
    
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            
        Select Case sts
            Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫ﾃﾞｰﾀ")
                Exit Function
        
        End Select
        
        '品番
        List = StrConv(ZAIKOREC.HIN_GAI, vbUnicode) & " "
        '倉庫
        List = List & StrConv(ZAIKOREC.Soko_No, vbUnicode) & " "
        '列
        List = List & StrConv(ZAIKOREC.Retu, vbUnicode) & " "
        '連
        List = List & StrConv(ZAIKOREC.Ren, vbUnicode) & " "
        '段
        List = List & StrConv(ZAIKOREC.Dan, vbUnicode) & " "
        '入荷日(年)
        List = List & Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & " "
        '入荷日(月)
        List = List & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & " "
        '入荷日(日)
        List = List & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
        '在庫数量
        wkQty = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
        If Len(wkQty) < 8 Then
            wkQty = Space(8 - Len(wkQty)) & wkQty
        End If
        List = List & wkQty & " "
        '仕入先
        List = List & StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode) & " "
        '仕入単価
        If IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
            wkTanka = Format(CDbl(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)), "#,##0.00")
            If Len(wkTanka) < 13 Then
                wkTanka = Space(13 - Len(wkTanka)) & wkTanka
            End If
        Else
            wkTanka = Space(13)
        End If
        List = List & wkTanka & " "
        
        
        List = List & "              " & StrConv(ZAIKOREC.HIN_GAI, vbUnicode) & _
                                            StrConv(ZAIKOREC.GOODS_ON, vbUnicode) & _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode) & _
                                            StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                            StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(ZAIKOREC.Dan, vbUnicode)
        
        
        
        List1(plstNYU).AddItem List
        
        com = BtOpGetNext
    Loop
        
    DoEvents

    List_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   前借ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

Dim wKey    As String

    Update_Proc = True
    
        
    
    
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Text1(ptxGOODS_ON).Text)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Text1(ptxNYUKA_YY).Text & Text1(ptxNYUKA_MM).Text & Text1(ptxNYUKA_DD).Text)
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Text1(ptxSoko_No))
    Call UniCode_Conv(K1_ZAIKO.Retu, Text1(ptxRetu))
    Call UniCode_Conv(K1_ZAIKO.Ren, Text1(ptxRen))
    Call UniCode_Conv(K1_ZAIKO.Dan, Text1(ptxDan))

    
    
    sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    
    Do
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Do
                End If
            
            Case BtErrKeyNotFound
                
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                
                Update_Proc = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    If Text1(ptxYUKO_Z_QTY).Locked = False Then
        
        If IsNumeric(Text1(ptxYUKO_Z_QTY).Text) Then
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(Text1(ptxYUKO_Z_QTY).Text), "00000000"))
        End If
    End If
    
    
    If IsNumeric(Text1(ptxSHIIRE_TANKA).Text) Then
        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, Format(CDbl(Text1(ptxSHIIRE_TANKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "00000000.00")
    End If
    
    Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, Text1(ptxSHIIRE_CODE).Text)
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    For i = ptxHIN_GAI To ptxSHIIRE_TANKA
        Text1(i).Text = ""
    Next i

    Text1(ptxYUKO_Z_QTY).Locked = True


    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).SetFocus
        List1(plstNYU).ListIndex = 0
    Else
        Text1(ptxHIN_GAI).SetFocus
    End If


    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   前借ﾃﾞｰﾀ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Text1(ptxGOODS_ON).Text)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Text1(ptxNYUKA_YY).Text & Text1(ptxNYUKA_MM).Text & Text1(ptxNYUKA_DD).Text)
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Text1(ptxSoko_No))
    Call UniCode_Conv(K1_ZAIKO.Retu, Text1(ptxRetu))
    Call UniCode_Conv(K1_ZAIKO.Ren, Text1(ptxRen))
    Call UniCode_Conv(K1_ZAIKO.Dan, Text1(ptxDan))

    
    
    sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    
    Do
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Do
                End If
            
            Case BtErrKeyNotFound
                
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                
                Delete_Proc = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "在庫ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    For i = ptxHIN_GAI To ptxSHIIRE_TANKA
        Text1(i).Text = ""
    Next i

    Text1(ptxYUKO_Z_QTY).Locked = True



    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).SetFocus
        List1(plstNYU).ListIndex = 0
    Else
        Text1(ptxHIN_GAI).SetFocus
    End If


    Delete_Proc = False


End Function

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
                    
        
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            End If
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Unload Me
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
            Command1(KeyCode - vbKeyF1).Value = True
    
        Case vbKeyEscape
    
            If Text1(ptxYUKO_Z_QTY).Locked = False Then
                Text1(ptxYUKO_Z_QTY).Locked = True
            Else
                Text1(ptxYUKO_Z_QTY).Locked = False
            End If
    
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
                                '資材前借ＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
   
    
    If List_Disp_Proc() Then
        Unload Me
    End If
    
    Show
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).ListIndex = 0
        List1(plstNYU).SetFocus
    End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                '資材前借ＯＰＥＮ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set Conv_ZAIKO1 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)

Dim W_KEY   As String
Dim sts     As Integer


    W_KEY = Right(List1(Index).List(List1(Index).ListIndex), 37)


    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Left(W_KEY, 20))
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Mid(W_KEY, 21, 1))
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Mid(W_KEY, 22, 8))
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Mid(W_KEY, 30, 2))
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(W_KEY, 32, 2))
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(W_KEY, 34, 2))
    Call UniCode_Conv(K1_ZAIKO.Dan, Mid(W_KEY, 36, 2))


    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            
            Text1(ptxHIN_GAI).Text = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            Text1(ptxSoko_No).Text = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            Text1(ptxRetu).Text = StrConv(ZAIKOREC.Retu, vbUnicode)
            Text1(ptxRen).Text = StrConv(ZAIKOREC.Ren, vbUnicode)
            Text1(ptxDan).Text = StrConv(ZAIKOREC.Dan, vbUnicode)

            Text1(ptxNYUKA_YY).Text = Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4)
            Text1(ptxNYUKA_MM).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2)
            Text1(ptxNYUKA_DD).Text = Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2)

            Text1(ptxYUKO_Z_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
        
            Text1(ptxSHIIRE_CODE).Text = Trim(StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
            If IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
                Text1(ptxSHIIRE_TANKA).Text = Format(CDbl((StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))), "#,##0.00")
            Else
                Text1(ptxSHIIRE_TANKA).Text = ""
            End If
                    
            Text1(ptxGOODS_ON).Text = StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
        
        Case BtErrKeyNotFound
        
            MsgBox "入力した項目はエラーです。"
            Text1(ptxHIN_GAI).SetFocus
            Exit Sub
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
            Exit Sub
    End Select

End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(plstNYU).ListCount > 0 And _
       List1(plstNYU).ListIndex < 0 Then
        List1(plstNYU).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim W_KEY   As String
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        
        W_KEY = Right(List1(Index).List(List1(Index).ListIndex), 37)
    
    
        Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Left(W_KEY, 20))
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Mid(W_KEY, 21, 1))
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, Mid(W_KEY, 22, 8))
        Call UniCode_Conv(K1_ZAIKO.Soko_No, Mid(W_KEY, 30, 2))
        Call UniCode_Conv(K1_ZAIKO.Retu, Mid(W_KEY, 32, 2))
        Call UniCode_Conv(K1_ZAIKO.Ren, Mid(W_KEY, 34, 2))
        Call UniCode_Conv(K1_ZAIKO.Dan, Mid(W_KEY, 36, 2))
    
    
        sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                
                Text1(ptxHIN_GAI).Text = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                Text1(ptxSoko_No).Text = StrConv(ZAIKOREC.Soko_No, vbUnicode)
                Text1(ptxRetu).Text = StrConv(ZAIKOREC.Retu, vbUnicode)
                Text1(ptxRen).Text = StrConv(ZAIKOREC.Ren, vbUnicode)
                Text1(ptxDan).Text = StrConv(ZAIKOREC.Dan, vbUnicode)
    
                Text1(ptxNYUKA_YY).Text = Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4)
                Text1(ptxNYUKA_MM).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2)
                Text1(ptxNYUKA_DD).Text = Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2)
    
                Text1(ptxYUKO_Z_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
            
                Text1(ptxSHIIRE_CODE).Text = Trim(StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                If IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
                    Text1(ptxSHIIRE_TANKA).Text = Format(CDbl((StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))), "#,##0.00")
                Else
                    Text1(ptxSHIIRE_TANKA).Text = ""
                End If
                        
                Text1(ptxGOODS_ON).Text = StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
            
            
            Case BtErrKeyNotFound
            
                MsgBox "入力した項目はエラーです。"
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
                Exit Sub
        End Select
    End If
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
        
        
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

