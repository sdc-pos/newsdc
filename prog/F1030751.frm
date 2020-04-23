VERSION 5.00
Begin VB.Form F1030751 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷予定登録"
   ClientHeight    =   6225
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   13350
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
   ScaleHeight     =   6225
   ScaleWidth      =   13350
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2835
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Width           =   345
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   6090
      MaxLength       =   1
      TabIndex        =   6
      Top             =   840
      Width           =   240
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   1665
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   21
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text 
      Height          =   360
      IMEMode         =   4  '全角ひらがな
      Index           =   19
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   20
      Top             =   3840
      Width           =   11490
   End
   Begin VB.TextBox Text 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   360
      IMEMode         =   4  '全角ひらがな
      Index           =   17
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2880
      Width           =   4875
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   645
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   22
      Top             =   1680
      Width           =   972
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   2865
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   3360
      Width           =   5145
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   9375
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   8655
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   7935
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   7215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   4935
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   4830
      MaxLength       =   7
      TabIndex        =   5
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3345
      MaxLength       =   2
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2625
      MaxLength       =   2
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   8925
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   4365
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   1725
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1680
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
      Left            =   9990
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   9150
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   8310
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   7470
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   6390
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   5550
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   4710
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   3870
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   2790
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   1950
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   1110
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5160
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
      Left            =   270
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   19
      Left            =   2625
      TabIndex        =   58
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID-№"
      Height          =   255
      Index           =   18
      Left            =   735
      TabIndex        =   57
      Top             =   360
      Width           =   645
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（元）"
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   56
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  名"
      Height          =   255
      Index           =   6
      Left            =   4725
      TabIndex        =   55
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   54
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "運送会社"
      Height          =   255
      Index           =   22
      Left            =   420
      TabIndex        =   53
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "備　　考"
      Height          =   255
      Index           =   21
      Left            =   420
      TabIndex        =   52
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "送 り 先"
      Height          =   255
      Index           =   20
      Left            =   420
      TabIndex        =   51
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   15
      Left            =   645
      TabIndex        =   50
      Top             =   1440
      Width           =   975
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
      Left            =   315
      TabIndex        =   49
      Top             =   5640
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "得 意 先"
      Height          =   255
      Index           =   17
      Left            =   420
      TabIndex        =   48
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   16
      Left            =   9135
      TabIndex        =   47
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   8415
      TabIndex        =   46
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   7695
      TabIndex        =   45
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ホスト棚番"
      Height          =   255
      Index           =   11
      Left            =   5745
      TabIndex        =   44
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ホスト倉庫"
      Height          =   255
      Index           =   10
      Left            =   3570
      TabIndex        =   43
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（先）"
      Height          =   255
      Index           =   9
      Left            =   2625
      TabIndex        =   42
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "予算単位"
      Height          =   255
      Index           =   8
      Left            =   525
      TabIndex        =   41
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票№"
      Height          =   255
      Index           =   4
      Left            =   3990
      TabIndex        =   40
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   3105
      TabIndex        =   39
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   2385
      TabIndex        =   38
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   37
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷予定数"
      Height          =   255
      Index           =   0
      Left            =   7665
      TabIndex        =   36
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  番"
      Height          =   255
      Index           =   5
      Left            =   2085
      TabIndex        =   35
      Top             =   1440
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030751"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const pcmbNAIGAI% = 0
Private Const pcmbMUKE_CODE% = 1
Private Const pcmbUNSOU_KAISHA% = 2





Private Const ptxMAX% = 19

Private Const ptxID_No% = 0
Private Const ptxID_SEQ% = 1



Private Const ptxYY% = 2
Private Const ptxMM% = 3
Private Const ptxDD% = 4

Private Const ptxNo% = 5
Private Const ptxSEQ% = 6

Private Const ptxCode% = 7
Private Const ptxName% = 8

Private Const ptxS_Qty% = 9

Private Const ptxMoto% = 10
Private Const ptxSaki% = 11

Private Const ptxSoko% = 12
Private Const ptxS_No% = 13
Private Const ptxRetu% = 14
Private Const ptxRen% = 15
Private Const ptxDan% = 16

Private Const ptxOKURISAKI% = 17
Private Const ptxMUKE_CODE% = 18
Private Const ptxBIKOU% = 19
                                   '画面初期状態を設定する
Private Sub Clear_Field(Optional Start_Pos As Integer = 0)
'----------------------------------------------------------------------------
'                   画面内容の消去
'----------------------------------------------------------------------------
Dim i As Integer
    
    For i = Start_Pos To ptxMAX
        Text(i).Text = ""
    Next i
    
End Sub
Private Function Item_Dsp() As Integer
'----------------------------------------------------------------------------
'                   品目マスタより各項目を表示する
'----------------------------------------------------------------------------

Dim sts As Integer


    Item_Dsp = True
                                                '国内外チェック
                                                'まず外部品番で読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxCode))
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxName) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text(ptxSoko) = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
            Text(ptxS_No) = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Text(ptxRetu) = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Text(ptxRen) = StrConv(ITEMREC.ST_REN, vbUnicode)
            Text(ptxDan) = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '内部品番で読み込み
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxCode).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    Text(ptxCode).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxName).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Text(ptxSoko).Text = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
                    Text(ptxS_No).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Text(ptxRetu).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Text(ptxRen).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Text(ptxDan).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                Case BtErrKeyNotFound
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Item_Dsp = SYS_ERR
                    Exit Function
            End Select
                
                
                
        Case Else
                
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Dsp = SYS_ERR
            Exit Function
        
    End Select
    
    Item_Dsp = False
    
End Function
Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------

Dim yn          As Integer
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim Flg         As Boolean
Dim Qty         As Long
Dim W_CYU_KBN   As String

    Err_Chk = True
                                        
                                        
                                        
    '-------------------------------------------    ID№
                                    '「０」以外が入力されたら登録済みチェック
    If Len(Text(ptxID_No).Text) = 0 Then
    Else
                                                '自動発番以外が入力されたら登録済みチェック
        If Not IsNumeric(Text(ptxID_No).Text) Then
        Else
            Text(ptxNo).Text = Format(CDbl(Text(ptxID_No).Text), "0000000")
        End If
        If Len(Text(ptxNo).Text) <> 7 Then
            Beep
            MsgBox "入力した項目はエラーです。（ID_NOは、7桁固定）"
            Text(ptxID_No).SetFocus
            Exit Function
        End If
        If Text(ptxID_SEQ).Text = "" Then
            Text(ptxID_SEQ).Text = "01"
        Else
            If Not IsNumeric(Text(ptxID_SEQ).Text) Then
                Text(ptxID_SEQ).Text = "01"
            End If
        End If
        Text(ptxID_SEQ).Text = Format(CInt(Text(ptxID_SEQ).Text), "00")
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Text(ptxID_No).Text & Text(ptxSEQ).Text)
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
                Beep
                MsgBox "入力した項目はエラーです。（出荷予定登録済み）"
                Text(ptxID_No).SetFocus
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
                                        
                                        
    '-------------------------------------------    伝票日付
    For i = ptxYY To ptxDD
        If Trim(Text(i)) = "" Then
            Text(i).Text = "0"
        End If
        
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです｡(伝票日付) "
            Text(i).SetFocus
            Exit Function
        Else
            RetBuf = Format(CLng(Text(i).Text), "0000")
            Text(i).Text = Right(RetBuf, Text(i).MaxLength)
        End If
    Next i
    If Not IsDate(Text(ptxYY).Text & "/" & Text(ptxMM).Text & "/" & Text(ptxDD).Text) Then
        Beep
        MsgBox "入力した項目はエラーです｡(伝票日付) "
        Text(i).SetFocus
        Exit Function
    End If
                                        
                                        
    '-------------------------------------------    伝票№
                                    '「０」以外が入力されたら登録済みチェック
    If Len(Text(ptxNo).Text) = 0 Then
    Else
                                                '自動発番以外が入力されたら登録済みチェック
        If Not IsNumeric(Text(ptxNo).Text) Then
        Else
            Text(ptxNo).Text = Format(CDbl(Text(ptxNo).Text), "0000000")
        End If
        If Len(Text(ptxNo).Text) <> 7 Then
            Beep
            MsgBox "入力した項目はエラーです。（伝票№は、7桁固定）"
            Text(ptxID_No).SetFocus
            Exit Function
        End If
        If Text(ptxSEQ).Text = "" Then
            Text(ptxSEQ).Text = "1"
        End If
        If Not IsNumeric(Text(ptxSEQ).Text) Then
            Text(ptxSEQ).Text = "1"
        End If
        
        
        Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text(ptxNo).Text)
        Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, Text(ptxSEQ).Text)
        
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        Select Case sts
            Case BtNoErr
                Beep
                MsgBox "入力した項目はエラーです。（出荷予定登録済み）"
                Text(ptxID_No).SetFocus
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
    '-------------------------------------------    品目
    sts = Item_Dsp
    Select Case sts
        Case False
        Case True
            Beep
            MsgBox "入力した項目はエラーです｡(品番) "
            Text(ptxCode).SetFocus
            Exit Function
        Case Else
            Err_Chk = SYS_ERR
            Exit Function
    End Select
    '-------------------------------------------    出荷予定数
    If Trim(Text(ptxS_Qty).Text) = "" Then
        Text(ptxS_Qty) = "0"
    End If
        
    If Not IsNumeric(Trim(Text(ptxS_Qty).Text)) Then
        Beep
        MsgBox "入力した項目はエラーです｡ (出荷予定数)"
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
    
    Text(ptxS_Qty).Text = Format(CLng(Text(ptxS_Qty).Text), "#0")
    If CLng(Text(ptxS_Qty).Text) <= 0 Then
        Beep
        MsgBox "入力した項目はエラーです｡ (出荷予定数)"
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
    '-------------------------------------------    得意先
    
    
    If Trim(Text(ptxMUKE_CODE).Text) = "" Then
        Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
        Call UniCode_Conv(K0_MTS.SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（得意先）"
                If MTS_Set_Proc() Then
                    Err_Chk = SYS_ERR
                    Exit Function
                End If
                Combo(pcmbMUKE_CODE).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
        Text(ptxMUKE_CODE).Text = Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8)
    Else
    
        Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        Call UniCode_Conv(K0_MTS.SS_CODE, "")
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
        
                    If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                        Combo(pcmbMUKE_CODE).ListIndex = i
                        Exit For
                    End If
                
        
                Next
            
            
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（得意先）"
                If MTS_Set_Proc() Then
                    Err_Chk = SYS_ERR
                    Exit Function
                End If
                Combo(pcmbMUKE_CODE).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    
    
    
    
    End If
    
    
    Err_Chk = False
    
    
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   出荷予定＆出荷予定(ホストイメージ)の追加
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer

Dim ID_NO   As String * 7
    
    
    
    Update_Proc = True
    
    
                                    'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If
    
    
    '-------------------------------------------    出荷予定
    
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                  '使用子機ＩＤ
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                  '使用中プログラム
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, "0")                                '完了区分
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                                 'データ種別
    Call UniCode_Conv(Y_SYUREC.JGYOBU, Last_JGYOBU)                         '事業部区分
    
                                                                            '注文区分
    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
    
                                                                            'ＩＤ№
    If Len(Trim(Text(ptxID_No).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Text(ptxID_No).Text & Text(ptxID_SEQ).Text)
        Call UniCode_Conv(Y_SYUREC.ID_NO, Text(ptxID_No).Text & Text(ptxID_SEQ).Text)
    Else
        sts = Den_No_Set_Proc(31, Last_JGYOBU, ID_NO)
        If sts Then
            Update_Proc = SYS_ERR
            GoTo Abort_Tran
        End If
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO & "01")
        Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO & "01")
        
    End If
                                                                            
    Call UniCode_Conv(Y_SYUREC.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))    '国内外
                                                                    
    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, Text(ptxCode).Text)              '品目番号
    Call UniCode_Conv(Y_SYUREC.HIN_NO, Text(ptxCode).Text)                  '品目番号
                                                                            '得意先コード
    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
                                                                            '直送先コード
    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                                                                            '出荷日
    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
                
    

    
    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                                  '事業場
    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                                'データ区分
    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                                '取引区分
                                                                            
    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                                                                            
    If Len(Trim(Text(ptxID_No).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.DEN_NO, Text(ptxNo).Text & Text(ptxSEQ).Text)
    Else                                                                    '伝票№
        sts = Den_No_Set_Proc(30, Last_JGYOBU, ID_NO)
        If sts Then
            Update_Proc = SYS_ERR
            GoTo Abort_Tran
        End If
        Call UniCode_Conv(Y_SYUREC.DEN_NO, ID_NO & "1")
    End If
                                                                            '出庫数量
    Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
        
    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                             '出庫収支
    
    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
    
    
    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                                 'オーダー番号
    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                                 'アイテム番号
    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                                                                            '得意先名称
    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
                                                                            '注文区分名称
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_2)
                                                                            '品名
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        
    
    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
        
        
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
        
                                                                            'ホスト棚番
    Call UniCode_Conv(Y_SYUREC.HTANABAN, Text(ptxS_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text)
    
    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                               '完了日付
    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                                 '完了日付
    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                              '検品日付
    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                                 '特売り区分
    
    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")                      '実績数量
    
                                                                            '更新日時
    Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")                       '検品担当者
    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")                              '検品時刻
                                                                            '上位ﾘﾝｸ用向け先ｺｰﾄﾞ
    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")                               '出荷実績連携№
    Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")                              '画面検品ﾌﾗｸﾞ
    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")                            '検品時巣量
    
    
    
    Call UniCode_Conv(Y_SYUREC.FILLER, "")

    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            Case BtErrDuplicates
                If Len(Trim(Text(ptxID_No).Text)) = 0 Then
                                            '自動発番データ重複は再発行
                    
                    ans = MsgBox("伝票IDが、重複しています。再発番しますか？", vbYesNo, "確認入力")
                    
                    If ans = vbYes Then
                    
                        sts = Den_No_Set_Proc(31, Last_JGYOBU, ID_NO)
                        If sts Then
                            Update_Proc = SYS_ERR
                            GoTo Abort_Tran
                        End If
        
                        Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO & "01")
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO & "01")
                    
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, ID_NO)
                    Else
                        Update_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                
                Else
                    ans = MsgBox("伝票IDが、重複しています。更新処理中止します", vbOK, "確認入力")
                    
                    Call File_Error(sts, BtOpInsert, "出荷予定データ")
                    Update_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "出荷予定データ")
                Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
    Loop
    
    '-------------------------------------------    出荷予定(ﾎｽﾄｲﾒｰｼﾞ)
    'ID_NO
    Call UniCode_Conv(Y_SYU_HREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
    '№
    Call UniCode_Conv(Y_SYU_HREC.SYUKA_NO, "")
    '出荷日
    Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, Text(ptxYY).Text & _
                                            Text(ptxMM).Text & _
                                            Text(ptxDD).Text)
    '送り先名
    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, Text(ptxOKURISAKI).Text)
    '売り伝
    Call UniCode_Conv(Y_SYU_HREC.URIDEN, "0")
    '伝票番号
    Call UniCode_Conv(Y_SYU_HREC.DEN_NO, Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 7))
    '追番
    Call UniCode_Conv(Y_SYU_HREC.SEQ_NO, Right(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)), 1))
    '品番
    Call UniCode_Conv(Y_SYU_HREC.HIN_NO, Text(ptxCode).Text)
    '数量
    Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
    '注文№
    Call UniCode_Conv(Y_SYU_HREC.ODER_NO, "")
    '得意先
    Call UniCode_Conv(Y_SYU_HREC.MUKE_CODE, Text(ptxMUKE_CODE).Text)
    '得意先名
    Call UniCode_Conv(Y_SYU_HREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
    '備考
    Call UniCode_Conv(Y_SYU_HREC.BIKOU, Trim(Text(ptxBIKOU).Text))
    '運送会社名
    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, Trim(Combo(pcmbUNSOU_KAISHA).Text))
    '取込み日時（入力日時）
    Call UniCode_Conv(Y_SYU_HREC.INS_NOW, Format(Now, "YYYYMMDDHHMMSS"))
    'ﾃﾞｰﾀ発生順
    Call UniCode_Conv(Y_SYU_HREC.DATA_CNT, "00001")
    '送り状№
    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
    '検品日時
    Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
    '検品担当者
    Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
    '口数
    Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "00")
    'FILLER
    Call UniCode_Conv(Y_SYU_HREC.FILLER, "")
    
    
    
    Do
        sts = BTRV(BtOpInsert, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            Case BtErrDuplicates
                'ﾃﾞｰﾀの矛盾が発生するのでこのまま処理中断
                ans = MsgBox("伝票IDが、重複しています。更新処理中止します", vbOK, "確認入力")
                
                Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄﾃﾞｰﾀ)データ")
                Update_Proc = SYS_CANCEL
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄﾃﾞｰﾀ)データ")
                Update_Proc = SYS_ERR
                GoTo Abort_Tran
        End Select
    Loop
                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        GoTo Abort_Tran
    End If
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If

    Beep
    MsgBox "伝票№:" & StrConv(Y_SYU_HREC.DEN_NO, vbUnicode) & "-" & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode) _
                & " ID:" & Left(StrConv(Y_SYUREC.ID_NO, vbUnicode), 7) & "-" & Right(Trim(StrConv(Y_SYUREC.ID_NO, vbUnicode)), 2)

    Update_Proc = False
    Exit Function
'異常終了
Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

End Function



Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
                                            '入力項目のクリアーとみなす
    Select Case Index
        Case pcmbNAIGAI
            Text(ptxCode).SetFocus
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE).Text = Trim(Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
            
            Text(ptxBIKOU).SetFocus
            
    End Select

End Sub

Private Sub Combo_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbNAIGAI
            Text(ptxCode).SetFocus
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE).Text = Trim(Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
            
            Text(ptxBIKOU).SetFocus
            
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0                      '更新
                                    'エラーチェック
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False, True
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
            
            Text(ptxID_SEQ).Text = ""
            Text(ptxSEQ).Text = ""
            Text(ptxID_No).SetFocus
                    
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
Dim i               As Integer
Dim c               As String * 128
Dim sts             As Integer
Dim UNSOU_KAISHA    As Variant

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
    LOG_F = RTrim(c)
                                
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030751.Caption = "大阪ＰＣ用　出荷予定登録（" + RTrim(JGYOBU_T(i).NAME) + ")"
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
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データファイルＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '画面初期設定
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & Space(5) & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & Space(5) & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
                        
                        
'運送会社
    Combo(pcmbUNSOU_KAISHA).Clear
                                '運送会社名称獲得
    If GetIni(App.EXEName, "UNSOU_KAISHA", "SYS", c) Then
    Else
        UNSOU_KAISHA = Split(Trim(c), ",", -1)
        For i = 0 To UBound(UNSOU_KAISHA)
            Combo(pcmbUNSOU_KAISHA).AddItem UNSOU_KAISHA(i)
        Next i
    End If
    Combo(pcmbUNSOU_KAISHA).ListIndex = 0
                        
                        '向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If
            
            
            
    Call Clear_Field
    Text(ptxYY) = Mid(Date, 1, 4)
    Text(ptxMM) = Mid(Date, 6, 2)
    Text(ptxDD) = ""
    
    Text(ptxID_No).SetFocus

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
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
                                            '出荷予定データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データファイル")
        End If
    End If
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データファイル")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030751 = Nothing

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
    F1030751.Caption = "大阪ＰＣ用　出荷予定登録（" + RTrim(JGYOBU_T(Index).NAME) + "）"
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
Dim RetBuf  As String
Dim i       As Integer
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Select Case Index
        Case ptxCode                '品目コード
            sts = Item_Dsp()
            Select Case sts
                Case False
                Case True
                    Beep
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(Index).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
        Case ptxMUKE_CODE         '向け先（コード入力用）
            Call UniCode_Conv(K0_MTS.MUKE_CODE, Text(Index).Text)
            Call UniCode_Conv(K0_MTS.SS_CODE, "")
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                                
                Case BtNoErr
                                
                Case BtErrKeyNotFound
                    Beep
                    MsgBox "入力した項目はエラーです。(向け先コード)"
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            
    End Select
    
    For i = Index + 1 To ptxMAX
        If Text(i).Visible And Text(i).Enabled And Not Text(i).Locked Then
            Text(i).SetFocus
            Exit Sub
        End If
    Next i
    Combo(pcmbUNSOU_KAISHA).SetFocus
    
End Sub
Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String


    MTS_Set_Proc = True
    
    com = BtOpGetFirst
    
    Combo(pcmbMUKE_CODE).Clear
    
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                MTS_Set_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & Space(8)
        Combo(pcmbMUKE_CODE).AddItem Edit
    
    
        com = BtOpGetNext
    Loop




    If Combo(pcmbMUKE_CODE).ListCount = 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If


    MTS_Set_Proc = False
End Function


