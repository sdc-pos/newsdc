VERSION 5.00
Begin VB.Form F1020401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出庫登録"
   ClientHeight    =   8250
   ClientLeft      =   2445
   ClientTop       =   3315
   ClientWidth     =   16260
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
   ScaleHeight     =   8250
   ScaleWidth      =   16260
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1920
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   60
      Top             =   720
      Width           =   1572
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   1920
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   59
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   300
      Left            =   5355
      Sorted          =   -1  'True
      TabIndex        =   58
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "数量"
      Height          =   1335
      Left            =   360
      TabIndex        =   45
      Top             =   4680
      Width           =   2775
      Begin VB.TextBox Text 
         Alignment       =   1  '右揃え
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   13
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text 
         Alignment       =   1  '右揃え
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   12
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "未商品"
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   47
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "商品化済"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   1  'ｵﾝ
      Index           =   14
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   15
      Top             =   6240
      Width           =   3735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   5100
      Index           =   0
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   7575
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   8
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3120
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
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "最  新"
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7320
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Height          =   375
      Index           =   2
      Left            =   11760
      TabIndex        =   54
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番(内部)"
      Height          =   255
      Index           =   22
      Left            =   11520
      TabIndex        =   67
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  棚番"
      Height          =   255
      Index           =   27
      Left            =   9840
      TabIndex        =   66
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入荷日"
      Height          =   255
      Index           =   28
      Left            =   8760
      TabIndex        =   65
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "在庫数"
      Height          =   255
      Index           =   29
      Left            =   14160
      TabIndex        =   64
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "商品化済(*)"
      Height          =   255
      Index           =   31
      Left            =   7800
      TabIndex        =   63
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "済(*)"
      Height          =   255
      Index           =   34
      Left            =   7800
      TabIndex        =   62
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "合計"
      Height          =   255
      Index           =   21
      Left            =   12120
      TabIndex        =   61
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblSelQty 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   4680
      TabIndex        =   57
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl_ZAN_QTY 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Height          =   375
      Left            =   6240
      TabIndex        =   56
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lbl_ZAN_T 
      BackColor       =   &H00FFFFFF&
      Caption         =   "指示残数"
      Height          =   255
      Left            =   5040
      TabIndex        =   55
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   11280
      TabIndex        =   53
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   51
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "＋"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   9480
      TabIndex        =   50
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   49
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "商品化済(*)"
      Height          =   255
      Index           =   17
      Left            =   8160
      TabIndex        =   48
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "総出庫指示数"
      Height          =   255
      Index           =   14
      Left            =   4560
      TabIndex        =   44
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblTanto_Name 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   2040
      TabIndex        =   43
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "担当"
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   42
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "メ　モ"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   41
      Top             =   6360
      Width           =   735
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
      TabIndex        =   40
      Top             =   7920
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "内外"
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "作業宣言"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   38
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（内部）"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   37
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   36
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   35
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入荷日"
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   34
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  名"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   33
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  番"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   32
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   31
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "－"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   29
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   28
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "未商品"
      Height          =   255
      Index           =   19
      Left            =   10080
      TabIndex        =   52
      Top             =   6240
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
Attribute VB_Name = "F1020401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim WS_NO   As String * 3

Dim MENU_NO As String * 2   '2007.07.11

Private Const ptxTanto_Code% = 0    '担当者コード
Private Const ptxHin_Gai% = 1       '品番(外部)
Private Const ptxTotal_Qty% = 2     '総出庫数
Private Const ptxHin_Name% = 3      '品名
Private Const ptxSoko_No% = 4       '倉庫№
Private Const ptxRetu% = 5          '列
Private Const ptxRen% = 6           '連
Private Const ptxDan% = 7           '段
Private Const ptxNyuka_DT_YY% = 8   '入荷日　年
Private Const ptxNyuka_DT_MM% = 9   '入荷日　月
Private Const ptxNyuka_DT_DD% = 10  '入荷日　日
Private Const ptxHin_Nai% = 11      '品番（内部）
Private Const ptxSumi_QTY% = 12     '数量（商品化済み）
Private Const ptxMi_QTY% = 13       '数量（未商品）
Private Const ptxMEMO% = 14         'メモ

Private Const Text_Max% = 14        '

Private Const pcmbSagyo% = 0        '作業宣言
Private Const pcmbNaigai% = 1       '内外

Private Const PlstZaiko% = 0        '在庫

Private Type MENU_TBL_Tag
    CODE    As String * 1
    NAME    As String * 4
    TYPE    As String * 1
    YOIN    As String * 1
End Type

Private MENU_TBL()  As MENU_TBL_Tag

'Private Const Last_Update_Day$ = "(F102040 2017.09.27 09:30)"
Private Const Last_Update_Day$ = "[F102040] 2019.07.11 12:00)"

Private Function List_Disp_Proc() As Integer

Dim NAIGAI          As String * 1
Dim sts             As Integer
Dim com             As Integer
Dim Soko_No         As String * 2

Dim Edit            As String
Dim RetBuf          As String

Dim GK_GOODS_ON     As Long
Dim GK_GOODS_OFF    As Long


    
    List_Disp_Proc = True
    
    List1(PlstZaiko).Clear
                                                    
    If Combo(pcmbNaigai).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                            '要因マスタより倉庫番号獲得
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Or _
        Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_OUT Then
                                    '要因マスタ読み込み
        Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Right(Combo(pcmbSagyo).Text, 2), 1))
        Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Combo(pcmbSagyo).Text, 1))
        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                Combo(pcmbSagyo).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
        
        Soko_No = StrConv(YOINREC.Soko_No, vbUnicode)
    End If
                                                    
    GK_GOODS_ON = 0
    GK_GOODS_OFF = 0
                                                    
                                                    '在庫データ読み込み
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)
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
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> RTrim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫データ")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
                                            
        If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_OUT And _
            Soko_No = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
        Else
                                            
                                            
                                                '棚マスタ読込み
            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, StrConv(ZAIKOREC.Retu, vbUnicode))
            Call UniCode_Conv(K0_TANA.Ren, StrConv(ZAIKOREC.Ren, vbUnicode))
            Call UniCode_Conv(K0_TANA.Dan, StrConv(ZAIKOREC.Dan, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound           '該当レコード無しは禁止扱い
                    Call UniCode_Conv(TANAREC.KAHI_KBN, KAHI_KBN_NG)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
        
            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
            Else
                If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
                    If Soko_No = StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                        If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                            GK_GOODS_ON = GK_GOODS_ON + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            Edit = "* "
                        Else
                            GK_GOODS_OFF = GK_GOODS_OFF + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            Edit = "  "
                        End If
                        
                        Edit = Edit & "     "       '2017.08.04
                                        
                        
                        Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
                        Edit = Edit & StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-" & _
                                        StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & _
                                        StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & _
                                        StrConv(ZAIKOREC.Dan, vbUnicode) & " "
                        Edit = Edit & StrConv(ZAIKOREC.HIN_NAI, vbUnicode) & " "
                        RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
                        If Len(Trim(RetBuf)) < 8 Then
                            RetBuf = Space(8 - Len(Trim(RetBuf))) & Trim(RetBuf)
                        End If
                        Edit = Edit & RetBuf
                        List1(PlstZaiko).AddItem Edit
                    End If
                Else
                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                        GK_GOODS_ON = GK_GOODS_ON + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        Edit = "* "
                    Else
                        GK_GOODS_OFF = GK_GOODS_OFF + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        Edit = "  "
                    End If
                    
                    Edit = Edit & "     "       '2017.08.04
                    
                    
                    Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                            Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                            Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
                    Edit = Edit & StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-" & _
                            StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & _
                            StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & _
                            StrConv(ZAIKOREC.Dan, vbUnicode) & " "
                    Edit = Edit & StrConv(ZAIKOREC.HIN_NAI, vbUnicode) & " "
                    RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
                    If Len(Trim(RetBuf)) < 8 Then
                        RetBuf = Space(8 - Len(Trim(RetBuf))) & Trim(RetBuf)
                    End If
                        
                    Edit = Edit & RetBuf
                    List1(PlstZaiko).AddItem Edit
                End If
            End If
        End If
        com = BtOpGetNext
    Loop
                                                
    lblTotal(0).Caption = Format(GK_GOODS_ON, "#0")
    lblTotal(1).Caption = Format(GK_GOODS_OFF, "#0")
    lblTotal(2).Caption = Format(GK_GOODS_ON + GK_GOODS_OFF, "#0")
                                                
                                                '在庫無し
    If List1(PlstZaiko).ListCount = 0 Then
        List_Disp_Proc = True
        Exit Function
    End If
    
    List_Disp_Proc = False

End Function

Private Function Err_Chk() As Integer
            
Dim YOIN            As String * 2
Dim NAIGAI          As String * 1

Dim i               As Integer
Dim sts             As Integer


    Err_Chk = True
                                    
                                    
    If Trim(Text(ptxTanto_Code).Text) = "" Then         '2016.04.20
        Beep                                            '2016.04.20
        MsgBox "担当者コードが入力されていません。"     '2016.04.20
        Text(ptxTanto_Code).SetFocus                    '2016.04.20
        Exit Function                                   '2016.04.20
    End If                                              '2016.04.20
                                    
                                    '担当者のチェック
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTanto_Code).Text)
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Beep
            MsgBox "入力した項目はエラーです｡ (担当者　未登録エラー)"
            Text(ptxTanto_Code).SetFocus
            Exit Function
        Case Else
           Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Err_Chk = SYS_ERR
            Exit Function
    End Select
                                    '作業宣言
    If Combo(pcmbSagyo).Text = "" Then
        Beep
        MsgBox "作業を選択してください。"
        Combo(pcmbSagyo).SetFocus
        Exit Function
    End If
    
    YOIN = Right(Combo(pcmbSagyo).Text, 2)
                                    '要因マスタ読み込み
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Right(Combo(pcmbSagyo).Text, 2), 1))
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Combo(pcmbSagyo).Text, 1))
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Beep
            MsgBox "入力した項目はエラーです｡ (要因　未登録エラー)"
            Combo(pcmbSagyo).SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "要因マスタ")
            Err_Chk = SYS_ERR
            Exit Function
    End Select
    
    
    If Trim(Text(ptxHin_Gai).Text) = "" Then            '2016.04.20
        Beep                                            '2016.04.20
        MsgBox "品番コードが入力されていません。"       '2016.04.20
        Text(ptxHin_Gai).SetFocus                       '2016.04.20
        Exit Function                                   '2016.04.20
    End If                                              '2016.04.20
    
    If Combo(pcmbNaigai).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
            
    sts = Item_Read_Proc()
    Select Case sts
        Case False
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                Text(ptxHin_Nai).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
            End If
        
        Case True
            Text(ptxHin_Name).Text = ""
            Beep
            MsgBox "入力した項目はエラーです｡ (品番　未登録エラー)"
            Text(ptxHin_Gai).SetFocus
            Exit Function
        Case SYS_ERR
            Err_Chk = SYS_ERR
            Exit Function
    End Select
                                                
    If YOIN = YOIN_FURIKAE Then
                                                '国内外振替時の振替後品番のチェック
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        If NAIGAI = NAIGAI_NAI Then
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
        Else
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        End If
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力したコードは登録されていません。（振替え品目）"
                Text(ptxHin_Gai).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
                                                
    For i = ptxSoko_No To ptxNyuka_DT_DD
        If Len(Trim(Text(i).Text)) = 0 Then
            Beep
            MsgBox "入力した項目はエラーです。（必須入力）"
            If Text(ptxSoko_No).Locked Then
                List1(PlstZaiko).SetFocus
                Exit Function
            Else
'2012.12.20                Text(ptxSoko_No).SetFocus
                Text(i).SetFocus                        '2012.12.20
                Exit Function
            End If
        Else
            If i = ptxSoko_No Then
                Text(i).Text = StrConv(Text(i).Text, vbUpperCase)       '2016.01.26
            Else
                If IsNumeric(Text(i).Text) Then
                    If i = ptxNyuka_DT_YY Then
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    Else
                        Text(i).Text = Format(CInt(Text(i).Text), "00")
                    End If
                Else
                    Beep
                    MsgBox "入力した項目はエラーです。（数値入力）"
                    If Text(ptxSoko_No).Locked Then
                        List1(PlstZaiko).SetFocus
                        Exit Function
                    Else
'2012.12.20                        Text(ptxSoko_No).SetFocus
                        Text(i).SetFocus                '2012.12.20
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
                                                
                                                '入庫処理のチェック
    If Mid(YOIN, 1, 1) = ACT_ZAITEI_IN Or _
        Mid(YOIN, 1, 1) = ACT_IDO_IN Then
                                                '倉庫混載チェック
        Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSoko_No).Text)
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力したコードは登録されていません。（倉庫マスタ）"
                Text(ptxSoko_No).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
                    
        If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
            If StrConv(SOKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                StrConv(SOKOREC.NAIGAI, vbUnicode) <> NAIGAI Then
                Beep
                MsgBox "入力した項目はエラーです（混載エラー）"
                Text(ptxSoko_No).SetFocus
                Exit Function
            End If
        End If
        
        For i = ptxRetu To ptxDan
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。（数値入力）"
                Text(i).SetFocus
                Exit Function
            Else
                Text(i).Text = Format(CInt(Text(i).Text), "00")
            End If
        Next i
                                            '棚マスタチェック
        Call UniCode_Conv(K0_TANA.Soko_No, Text(ptxSoko_No).Text)
        Call UniCode_Conv(K0_TANA.Retu, Text(ptxRetu).Text)
        Call UniCode_Conv(K0_TANA.Ren, Text(ptxRen).Text)
        Call UniCode_Conv(K0_TANA.Dan, Text(ptxDan).Text)
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力したコードは登録されていません。（棚マスタ）"
                Text(ptxSoko_No).SetFocus
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
                                            '棚状態のチェック
        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
            Beep
            MsgBox "入力した項目はエラーです（棚使用不可）"
            Text(ptxSoko_No).SetFocus
            Exit Function
        End If
                                            '入荷日付のチェック
        If Not IsNumeric(Text(ptxNyuka_DT_MM).Text) Then
            Beep
            MsgBox "入力した項目はエラーです（数値入力）"
            Text(ptxNyuka_DT_MM).SetFocus
            Exit Function
        Else
            Text(ptxNyuka_DT_MM).Text = Format(CInt(Text(ptxNyuka_DT_MM).Text), "00")
        End If
        If Not IsNumeric(Text(ptxNyuka_DT_DD).Text) Then
            Beep
            MsgBox "入力した項目はエラーです（数値入力）"
            Text(ptxNyuka_DT_DD).SetFocus
            Exit Function
        Else
            Text(ptxNyuka_DT_DD).Text = Format(CInt(Text(ptxNyuka_DT_DD).Text), "00")
        End If
    
        If Not IsDate(Text(ptxNyuka_DT_YY).Text & "/" & Text(ptxNyuka_DT_MM).Text & "/" & Text(ptxNyuka_DT_DD).Text) Then
            Beep
            MsgBox "入力した項目はエラーです（日付入力）"
            Text(ptxNyuka_DT_YY).SetFocus
            Exit Function
        End If
    End If
                                                '数値項目
    If Text(ptxSumi_QTY).Text = "" Then
        Text(ptxSumi_QTY).Text = "0"
    End If
    
    If Not IsNumeric(Text(ptxSumi_QTY).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。（数値入力）"
        Text(ptxSumi_QTY).SetFocus
        Exit Function
    
    Else
        Text(ptxSumi_QTY).Text = Format(CLng(Text(ptxSumi_QTY).Text), "#0")
        If CLng(Text(ptxSumi_QTY).Text) < 0 Then
            Beep
            MsgBox "入力した項目はエラーです。（整数入力）"
            Text(ptxSumi_QTY).SetFocus
            Exit Function
        End If
    End If

    If Text(ptxMi_QTY).Text = "" Then
        Text(ptxMi_QTY).Text = "0"
    End If
    
    If Not IsNumeric(Text(ptxMi_QTY).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。（数値入力）"
        Text(ptxMi_QTY).SetFocus
        Exit Function
    Else
        Text(ptxMi_QTY).Text = Format(CLng(Text(ptxMi_QTY).Text), "#0")
        If CLng(Text(ptxMi_QTY).Text) < 0 Then
            Beep
            MsgBox "入力した項目はエラーです。（整数入力）"
            Text(ptxMi_QTY).SetFocus
            Exit Function
        End If
    End If

    If CLng(Text(ptxSumi_QTY).Text) = 0 And CLng(Text(ptxMi_QTY).Text) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。（必須入力）"
        Text(ptxSumi_QTY).SetFocus
        Exit Function
    End If
    
    
    If Mid(YOIN, 1, 1) = ACT_ZAITEI_IN Then
    Else
        If (CLng(Text(ptxSumi_QTY).Text) + CLng(Text(ptxMi_QTY).Text)) > CLng(lblSelQty) Then
            Beep
            MsgBox "入力した項目はエラーです。（数量オーバー）"
        
            If Text(ptxSumi_QTY).Locked Then
                Text(ptxMi_QTY).SetFocus
                Exit Function
            Else
                Text(ptxSumi_QTY).SetFocus
                Exit Function
            End If
    
        End If
    
        If Mid(YOIN, 1, 1) = ACT_IDO_IN Then
        Else
            If (CLng(Text(ptxSumi_QTY).Text) + CLng(Text(ptxMi_QTY).Text)) > CLng(lbl_ZAN_QTY.Caption) Then
                Beep
                MsgBox "入力した項目はエラーです。（数量オーバー）"
                If Text(ptxSumi_QTY).Locked Then
                    Text(ptxMi_QTY).SetFocus
                    Exit Function
                Else
                    Text(ptxSumi_QTY).SetFocus
                    Exit Function
                End If
        
            End If
        End If
    
    End If
    
    Err_Chk = False
End Function
Private Sub Zaiko_Detail_Proc()
        
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
                                        '移動入庫時は行き先を入力の為、初期表示なし
        Text(ptxSoko_No).Text = ""
        Text(ptxRetu).Text = ""
        Text(ptxRen).Text = ""
        Text(ptxDan).Text = ""
    Else
        
        Text(ptxSoko_No).Text = StrConv(ZAIKOREC.Soko_No, vbUnicode)
        Text(ptxRetu).Text = StrConv(ZAIKOREC.Retu, vbUnicode)
        Text(ptxRen).Text = StrConv(ZAIKOREC.Ren, vbUnicode)
        Text(ptxDan).Text = StrConv(ZAIKOREC.Dan, vbUnicode)
    
    End If
    
    Text(ptxNyuka_DT_YY).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4)
    Text(ptxNyuka_DT_MM).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2)
    Text(ptxNyuka_DT_DD).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2)
    Text(ptxHin_Nai).Text = StrConv(ZAIKOREC.HIN_NAI, vbUnicode)
    
    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
        
        If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
            Text(ptxSumi_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
        Else
            If CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) <= CLng(lbl_ZAN_QTY.Caption) Then
                Text(ptxSumi_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
            Else
                Text(ptxSumi_QTY).Text = Format(CLng(lbl_ZAN_QTY.Caption), "#0")
            End If
        End If
    
        Text(ptxMi_QTY).Text = ""
    
    Else
        If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
            Text(ptxMi_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
        Else
            If Not IsNumeric(lbl_ZAN_QTY.Caption) Then
                 lbl_ZAN_QTY.Caption = Text(ptxTotal_Qty).Text
            End If
        
            If CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) <= CLng(lbl_ZAN_QTY.Caption) Then
                Text(ptxMi_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
            Else
                Text(ptxMi_QTY).Text = Format(CLng(lbl_ZAN_QTY.Caption), "#0")
            End If
        End If
    
        Text(ptxSumi_QTY).Text = ""
    
    End If

    lblSelQty = StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)

End Sub

                                            
Private Function Update_Proc() As Integer

Dim YOIN            As String * 2
Dim NAIGAI          As String * 1

Dim TO_NAIGAI       As String * 1

Dim sts             As Integer

Dim IDO_Soko_No     As String * 2

Dim WK_CODE         As String * 5       '2007.05.28
Dim WK_TANKA        As String * 11      '2007.05.28



    Update_Proc = True
    
    Call Input_Lock
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If



    YOIN = Right(Combo(pcmbSagyo).Text, 2)      '要因
    
    If Combo(pcmbNaigai).Text = NAIGAI1 Then    '国内外の判定
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                                    
    
    Select Case Left(YOIN, 1)
        Case ACT_ZAITEI_IN                      '入庫
            If Last_JGYOBU = SHIZAI Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "入力した品番は登録されていません。（振替え品目）"
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
            
                '品目ﾏｽﾀの最新仕入先／単価が設定されていた時は、こちらの項目を使用  2007.05.28
                If Not IsNumeric(StrConv(ITEMREC.LAST_TANKA, vbUnicode)) Then
                    
                    WK_CODE = StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode)
                    WK_TANKA = StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)
                Else
                    WK_CODE = StrConv(ITEMREC.LAST_CODE, vbUnicode)
                    WK_TANKA = StrConv(ITEMREC.LAST_TANKA, vbUnicode)
                
                End If
            
            
            
                sts = Nyuko_Update_Proc(Last_JGYOBU, _
                                        NAIGAI, _
                                        Text(ptxHin_Gai).Text, _
                                        (Text(ptxNyuka_DT_YY).Text & Text(ptxNyuka_DT_MM).Text & Text(ptxNyuka_DT_DD).Text), _
                                        (Text(ptxSoko_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text), _
                                        YOIN, _
                                        CLng(Text(ptxSumi_QTY).Text), _
                                        CLng(Text(ptxMi_QTY).Text), _
                                        WS_NO, _
                                        Text(ptxTanto_Code).Text, , _
                                        Text(ptxMEMO).Text, _
                                        WK_CODE, _
                                        WK_TANKA)
                Select Case sts
                    Case False
                    Case Else
                        Update_Proc = sts
                        GoTo Abort_Tran
                End Select
            
            
            Else
            
            
            
                sts = Nyuko_Update_Proc(Last_JGYOBU, _
                                        NAIGAI, _
                                        Text(ptxHin_Gai).Text, _
                                        (Text(ptxNyuka_DT_YY).Text & Text(ptxNyuka_DT_MM).Text & Text(ptxNyuka_DT_DD).Text), _
                                        (Text(ptxSoko_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text), _
                                        YOIN, _
                                        CLng(Text(ptxSumi_QTY).Text), _
                                        CLng(Text(ptxMi_QTY).Text), _
                                        WS_NO, _
                                        Text(ptxTanto_Code).Text, , _
                                        Text(ptxMEMO).Text, , , , MENU_NO)  '2007.07.11 MENU_NO追加
                Select Case sts
                    Case False
                    Case Else
                        Update_Proc = sts
                        GoTo Abort_Tran
                End Select
            End If
        
        Case ACT_ZAITEI_OUT                     '出庫
            sts = Syuko_Update_Proc(Last_JGYOBU, _
                                    NAIGAI, _
                                    Text(ptxHin_Gai).Text, _
                                    (Text(ptxNyuka_DT_YY).Text & Text(ptxNyuka_DT_MM).Text & Text(ptxNyuka_DT_DD).Text), _
                                    (Text(ptxSoko_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text), _
                                    YOIN, _
                                    CLng(Text(ptxSumi_QTY).Text), _
                                    CLng(Text(ptxMi_QTY).Text), _
                                    0, _
                                    WS_NO, _
                                    Text(ptxTanto_Code).Text, , _
                                    Text(ptxMEMO).Text, , , , , , MENU_NO)  '2007.07.11 MENU_NO追加
            Select Case sts
                Case False
                Case Else
                    Update_Proc = sts
                    GoTo Abort_Tran
            End Select
        Case ACT_IDO_IN                     '移動入庫
                                        '移動対象倉庫設定
            IDO_Soko_No = StrConv(YOINREC.Soko_No, vbUnicode)

            sts = IDO_Update_Proc(Last_JGYOBU, _
                                    NAIGAI, _
                                    Text(ptxHin_Gai).Text, _
                                    (Text(ptxNyuka_DT_YY).Text & Text(ptxNyuka_DT_MM).Text & Text(ptxNyuka_DT_DD).Text), _
                                    (IDO_Soko_No & "01" & "01" & "01"), _
                                    (Text(ptxSoko_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text), _
                                    YOIN, _
                                    CLng(Text(ptxSumi_QTY).Text), _
                                    CLng(Text(ptxMi_QTY).Text), _
                                    WS_NO, _
                                    Text(ptxTanto_Code).Text, , _
                                    Text(ptxMEMO).Text, MENU_NO)    '2007.07.11 MENU_NO追加
            Select Case sts
                Case False
                Case Else
                    Update_Proc = sts
                    GoTo Abort_Tran
            End Select
                
                
        Case ACT_IDO_OUT                    '移動出庫
                                        '移動対象倉庫設定
            IDO_Soko_No = StrConv(YOINREC.Soko_No, vbUnicode)

            sts = IDO_Update_Proc(Last_JGYOBU, _
                                    NAIGAI, _
                                    Text(ptxHin_Gai).Text, _
                                    (Text(ptxNyuka_DT_YY).Text & Text(ptxNyuka_DT_MM).Text & Text(ptxNyuka_DT_DD).Text), _
                                    (Text(ptxSoko_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text), _
                                    (IDO_Soko_No & "01" & "01" & "01"), _
                                    YOIN, _
                                    CLng(Text(ptxSumi_QTY).Text), _
                                    CLng(Text(ptxMi_QTY).Text), _
                                    WS_NO, _
                                    Text(ptxTanto_Code).Text, , _
                                    Text(ptxMEMO).Text, MENU_NO)    '2007.07.11 MENU_NO追加
            Select Case sts
                Case False
                Case Else
                    Update_Proc = sts
                    GoTo Abort_Tran
            End Select
    End Select

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock

End Function



Private Sub Combo_Click(Index As Integer)

Dim sts     As Integer
Dim NAIGAI  As String * 1
'----------------------------------------------------------------------------
'                   コンボボックス入力（Click）処理
'----------------------------------------------------------------------------
    If Combo(pcmbSagyo).Text = "" Then Exit Sub
        
        
            
            
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        Unload Me
    End If
                
    Call Clear_field(1)
            
    Call Input_Change_Proc
    
            
    If Combo(pcmbNaigai).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                        '要因マスタ読み込み
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Right(Combo(pcmbSagyo).Text, 2), 1))
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Combo(pcmbSagyo).Text, 1))
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Beep
            MsgBox "入力した項目はエラーです｡ (未登録エラー)"
            Combo(Index).SetFocus
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "要因マスタ")
            Unload Me
    End Select
                
End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts     As Integer
Dim NAIGAI  As String * 1
'----------------------------------------------------------------------------
'                   コンボボックス入力（ＫｅｙＤｏｗｎ）処理
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbSagyo
                
'            Call Input_Lock
            
            If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
                Unload Me
            End If
                
            Call Clear_field(1)
'            Call Input_UnLock
            
            Call Input_Change_Proc
            
'2016.04.20            Combo(pcmbNaigai).SetFocus
'2016.04.20        Case pcmbNaigai
            If Combo(Index).Text = NAIGAI1 Then
                NAIGAI = NAIGAI_NAI
            Else
                NAIGAI = NAIGAI_GAI
            End If
                                        '要因マスタ読み込み
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Right(Combo(pcmbSagyo).Text, 2), 1))
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Combo(pcmbSagyo).Text, 1))
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Beep
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Combo(Index).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                    Unload Me
            End Select
                
            Text(ptxHin_Gai).SetFocus
    End Select
End Sub



Private Sub Command_Click(Index As Integer)
Dim yn      As Integer
Dim sts     As Integer


    Select Case Index
        Case 0
                                            'エラーチェック
'            Call Input_Lock
            
            Text(ptxHin_Gai).Text = RTrim(StrConv(Text(ptxHin_Gai).Text, vbUpperCase))
            
            
            sts = Err_Chk()
            
'            Call Input_UnLock
            
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
                    Case False, SYS_CANCEL, True
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
                        
            If lbl_ZAN_QTY.Visible Then
                lbl_ZAN_QTY.Caption = Format(CLng(lbl_ZAN_QTY.Caption) - (CLng(Text(ptxSumi_QTY).Text) + CLng(Text(ptxMi_QTY).Text)), "#0")
                If CLng(lbl_ZAN_QTY.Caption) <= 0 Then
                    Text(ptxTotal_Qty).Locked = False
                    Text(ptxTotal_Qty).Text = ""
                    lbl_ZAN_QTY.Caption = ""
            
                Else
                    sts = List_Disp_Proc()
                            
                    Select Case sts
                        Case False
                        Case True
                            If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                            Else
                                MsgBox "入力した品番には出庫可能な在庫が有りません。"
                                Text(ptxHin_Gai).SetFocus
                                Exit Sub
                            End If
                        Case Else
                            Unload Me
                    End Select
                            
                            
                    If List1(PlstZaiko).ListCount <= 0 Then
                        Text(ptxTotal_Qty).Locked = False
                        Text(ptxTotal_Qty).Text = ""
                        lbl_ZAN_QTY.Caption = ""
                    Else
                        List1(PlstZaiko).SetFocus
                        List1(PlstZaiko).ListIndex = 0
                        Call Clear_field(ptxSoko_No, 1)
                        Exit Sub
                    End If
                End If
            End If
            
            Call Clear_field(1, 0)
            
            Text(ptxHin_Gai).SetFocus
                
        Case 7                          '最新表示
            sts = Item_Read_Proc()
            Select Case sts
                Case False
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                        Text(ptxHin_Nai).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                    End If
                Case True
                    Text(ptxHin_Name).Text = ""
                    Beep
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(ptxHin_Gai).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
    
            sts = List_Disp_Proc()
            Select Case sts
                Case False
                Case True
                    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                    Else
                        Beep
                        MsgBox "入力した品番には出庫可能な在庫が有りません。"
                        Text(ptxHin_Gai).SetFocus
                        Exit Sub
                    End If
                Case Else
                    Unload Me
            End Select
            
            If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
            Else
                List1(PlstZaiko).ListIndex = 0
                List1(PlstZaiko).SetFocus
                Exit Sub
            End If
        Case 11                         '終了
                                        '在庫データ使用中解除
            If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
                Unload Me
            End If
            
            Unload Me
        Case Else
            Beep
    End Select
End Sub



Private Sub Form_DblClick()
'    PrintForm              '2017.07.22
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
    
Dim sBuffer As String * 255
Dim com     As String
    
'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If
    
    Show
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

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            Me.Caption = "入出庫登録（" + RTrim(JGYOBU_T(i).NAME) + ") " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

    Unload SubMenu(i)

'仮想倉庫番号番号取り込み
'    If Kaso_Soko_No_Set() Then
'        Beep
'        MsgBox "仮想倉庫の獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
                                'システム予約済要因取り込み
'    If SYSTEM_YOIN_Set() Then
'        Beep
'        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
'                                        '「前借り入荷」の要因
    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_MAEGARI = Trim(c)
                                        '「国内外振替え」の要因
    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_FURIKAE = Trim(c)
                                        '「国内外振替え」の要因
    If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_IN] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_FURIKAE_IN = Trim(c)
                                        '「国内外振替え」の要因
    If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_OUT] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_FURIKAE_OUT = Trim(c)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>    SYS.INI --> F102040.INI 2016.01.26
                                        'ﾒﾆｭｰ№獲得 2007.07.11
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = Trim(c)
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>    SYS.INI --> F102040.INI 2016.01.26



'端末番号取り込み
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタワークＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ（移動処理用）
    If wZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷実績（前借り）ＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '資材入荷実績（前借り）ＯＰＥＮ
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '作業実績ﾛｸﾞＯＰＥＮ
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    Combo(pcmbNaigai).AddItem NAIGAI1
    Combo(pcmbNaigai).AddItem NAIGAI2
    Combo(pcmbNaigai).ListIndex = 0
                                
                                '作業情報取り込み
    '作業タイプ設定
    i = 0
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve MENU_TBL(i)
        MENU_TBL(i).CODE = Trim(c)
        
        If GetIni("ACTION", "ACTION_NM" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_NM" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).NAME = Trim(c)
        
        If GetIni("ACTION", "ACTION_TYPE" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_TYPE" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).TYPE = Trim(c)
        
        If GetIni("ACTION", "ACTION_YOIN" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[ACTION]" & "[ACTION_YOIN" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).YOIN = Trim(c)
        i = i + 1
    Loop
                                
                                
                                '要因設定
    If Yoin_Set_Proc() Then
        Unload Me
    End If
        
    Combo(pcmbSagyo).ListIndex = 0
        
    Call Input_Change_Proc
        
        
        
        

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            
                                        '在庫データ使用中解除
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
    End If
                                            
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '品目マスタワークＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '在庫データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫データファイルＣＬＯＳＥ（移動処理用）
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '入荷実績（前借り）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷実績（前借り）")
        End If
    End If
    
                                            '資材入荷実績（前借り）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材入荷実績（前借り）")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1020401 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)

Dim sts         As Integer
Dim Location    As String * 8
Dim NYUKA_YMD   As String * 8
Dim NAIGAI      As String * 1
Dim End_Flg     As Boolean

Dim GOODS_FLG   As String * 1



                                        '入庫処理はリストボックスから選択不可
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
        Text(ptxSoko_No).SetFocus
        Exit Sub
    End If
                                                
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_OUT Or _
        Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_OUT Then
        
        If Not IsNumeric(Text(ptxTotal_Qty).Text) Then
            Exit Sub
        End If
    End If
                                                
                                                
                                                
    Call Input_Lock
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If
                                        '在庫データ使用中解除
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        End_Flg = True
        GoTo Abort_Tran
    End If
                                        
                                        
'>>>>>>> 2017.08.04
                                        'ロケーション獲得
'    Location = Mid(List1(Index).List(List1(Index).ListIndex), 14, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 17, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 20, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 23, 2)
                                        
                                        
                                        
    Location = Mid(List1(Index).List(List1(Index).ListIndex), 19, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 22, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 25, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 28, 2)
'>>>>>>> 2017.08.04
                                        
                                        
                                        
                                        
                                        '国内外獲得
    If Combo(pcmbNaigai).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
    End_Flg = False
    sts = Zaiko_Lock_Proc(Location, Last_JGYOBU, NAIGAI, Text(ptxHin_Gai).Text, WS_NO)
    Select Case sts
        Case False
        Case True, SYS_CANCEL
            GoTo Abort_Tran
        Case SYS_ERR
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                                
'>>>>>>> 2017.08.04
'    NYUKA_YMD = Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 3, 4) & _
'                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 8, 2) & _
'                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 11, 2)
    
    
    NYUKA_YMD = Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 8, 4) & _
                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 13, 2) & _
                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 16, 2)
'>>>>>>> 2017.08.04
    
    
    
    If Left(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 1) = "*" Then
        GOODS_FLG = GOODS_ON
    Else
        GOODS_FLG = GOODS_OFF
    End If
                                                '在庫データファイル読み込み
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)             '事業部
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)                  '国内外
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)  '品番（外部）
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_FLG)             '商品／未商品
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, NYUKA_YMD)             '入荷日
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Mid(Location, 1, 2))    '棚番　倉庫
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(Location, 3, 2))       '      列
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(Location, 5, 2))        '      連
    Call UniCode_Conv(K1_ZAIKO.Dan, Mid(Location, 7, 2))        '      段
        
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            Call Zaiko_Detail_Proc
        Case BtErrKeyNotFound
            Beep
            MsgBox "データ内容が変更されています。「最新」表示を選択してください。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "在庫データ")
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                        'トランザクション終了
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        End_Flg = True
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock

    If GOODS_FLG = GOODS_ON Then
        Text(ptxSumi_QTY).TabStop = True
        Text(ptxSumi_QTY).Locked = False
        Text(ptxMi_QTY).TabStop = False
        Text(ptxMi_QTY).Locked = True
    Else
        Text(ptxSumi_QTY).TabStop = False
        Text(ptxSumi_QTY).Locked = True
        Text(ptxMi_QTY).TabStop = True
        Text(ptxMi_QTY).Locked = False
    End If

    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
        Text(ptxSoko_No).SetFocus       '移動入庫は倉庫にフォーカス
    Else
        If GOODS_FLG = GOODS_ON Then
            Text(ptxSumi_QTY).SetFocus  '上記以外は数量にフォーカス
        Else
            Text(ptxMi_QTY).SetFocus
        End If
    End If
    
    Exit Sub

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        Unload Me
    End If

    If End_Flg Then
        Unload Me
    End If

    List1(PlstZaiko).ListIndex = 0
    List1(PlstZaiko).SetFocus
End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts         As Integer
Dim Location    As String * 8
Dim NYUKA_YMD   As String * 8
Dim NAIGAI      As String * 1
Dim End_Flg     As Boolean

Dim GOODS_FLG   As String * 1

                                        
    If List1(Index).ListCount = 0 Then
        Exit Sub
    End If
                                        
    If KeyCode <> vbKeyReturn Then Exit Sub

                                        '入庫処理はリストボックスから選択不可
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
        Text(ptxSoko_No).SetFocus
        Exit Sub
    End If
                                                
    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_OUT Or _
        Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_OUT Then
        
        If Not IsNumeric(Text(ptxTotal_Qty).Text) Then
            Exit Sub
        End If
    End If
                                                
                                                
                                                
    Call Input_Lock
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If
                                        '在庫データ使用中解除
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        End_Flg = True
        GoTo Abort_Tran
    End If
'>>>>>>> 2017.08.04
                                        'ロケーション獲得
'    Location = Mid(List1(Index).List(List1(Index).ListIndex), 14, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 17, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 20, 2) & _
'                Mid(List1(Index).List(List1(Index).ListIndex), 23, 2)
                                        
                                        
                                        
    Location = Mid(List1(Index).List(List1(Index).ListIndex), 19, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 22, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 25, 2) & _
                Mid(List1(Index).List(List1(Index).ListIndex), 28, 2)
'>>>>>>> 2017.08.04
                                        
                                        '国内外獲得
    If Combo(pcmbNaigai).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
    End_Flg = False
    sts = Zaiko_Lock_Proc(Location, Last_JGYOBU, NAIGAI, Text(ptxHin_Gai).Text, WS_NO)
    Select Case sts
        Case False
        Case True, SYS_CANCEL
            GoTo Abort_Tran
        Case SYS_ERR
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                                
'>>>>>>> 2017.08.04
'    NYUKA_YMD = Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 3, 4) & _
'                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 8, 2) & _
'                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 11, 2)
    
    
    NYUKA_YMD = Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 8, 4) & _
                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 13, 2) & _
                                            Mid(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 16, 2)
'>>>>>>> 2017.08.04
    
    If Left(List1(PlstZaiko).List(List1(PlstZaiko).ListIndex), 1) = "*" Then
        GOODS_FLG = GOODS_ON
    Else
        GOODS_FLG = GOODS_OFF
    End If
                                                '在庫データファイル読み込み
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)             '事業部
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)                  '国内外
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)  '品番（外部）
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_FLG)             '商品／未商品
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, NYUKA_YMD)             '入荷日
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Mid(Location, 1, 2))    '棚番　倉庫
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(Location, 3, 2))       '      列
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(Location, 5, 2))        '      連
    Call UniCode_Conv(K1_ZAIKO.Dan, Mid(Location, 7, 2))        '      段
        
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            Call Zaiko_Detail_Proc
        Case BtErrKeyNotFound
            Beep
            MsgBox "データ内容が変更されています。「最新」表示を選択してください。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "在庫データ")
            End_Flg = True
            GoTo Abort_Tran
    End Select
                                        'トランザクション終了
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        End_Flg = True
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock

    If GOODS_FLG = GOODS_ON Then
        Text(ptxSumi_QTY).TabStop = True
        Text(ptxSumi_QTY).Locked = False
        Text(ptxMi_QTY).TabStop = False
        Text(ptxMi_QTY).Locked = True
    Else
        Text(ptxSumi_QTY).TabStop = False
        Text(ptxSumi_QTY).Locked = True
        Text(ptxMi_QTY).TabStop = True
        Text(ptxMi_QTY).Locked = False
    End If

    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
        Text(ptxSoko_No).SetFocus       '移動入庫は倉庫にフォーカス
    Else
        If GOODS_FLG = GOODS_ON Then
            Text(ptxSumi_QTY).SetFocus  '上記以外は数量にフォーカス
        Else
            Text(ptxMi_QTY).SetFocus
        End If
    End If
    
    Exit Sub

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        Unload Me
    End If

    If End_Flg Then
        Unload Me
    End If

    List1(PlstZaiko).ListIndex = 0
    List1(PlstZaiko).SetFocus

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
    F1020401.Caption = "入出庫登録（" + RTrim(JGYOBU_T(Index).NAME) + ") " & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
                                '要因設定
    If Yoin_Set_Proc() Then
        Unload Me
    End If
    
    Combo(pcmbSagyo).ListIndex = 0
    
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop Then
        
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


    Select Case Index
    
        Case ptxTotal_Qty   '総出庫指示数
            
            If Trim(lbl_ZAN_QTY.Caption) = Trim(Text(ptxTotal_Qty).Text) Then
            
                Text(ptxTotal_Qty).Locked = False    '以降入力不可
            
            End If
            
    
    End Select
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
Dim sts As Integer
    
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case ptxTanto_Code                      '担当者ｺｰﾄﾞのチェック
            
            
            If Trim(Text(Index).Text) = "" Then                 '2016.04.20
                Beep                                            '2016.04.20
                MsgBox "担当者が入力されていません。"           '2016.04.20
                Text(ptxTanto_Code).SetFocus                    '2016.04.20
                Exit Sub                                        '2016.04.20
            End If                                              '2016.04.20
            
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTanto_Code).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                        Combo(pcmbSagyo).SetFocus
                        Exit Sub
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                        Text(ptxTanto_Code).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Unload Me
                End Select
                
        Case ptxHin_Gai                         '品番ｺｰﾄﾞのチェック
            Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
            
            
            If Trim(Text(Index).Text) = "" Then             '2016.04.20
                Beep                                        '2016.04.20
                MsgBox "品番が入力されていません。"         '2016.04.20
                Text(ptxHin_Gai).SetFocus                   '2016.04.20
                Exit Sub                                    '2016.04.20
            End If                                          '2016.04.20
            
            sts = Item_Read_Proc()
            Select Case sts
                Case False
                                                
    
    
                                                
                                                'ここまで戻ったら総数入力解除
                    If Trim(Text(ptxHin_Name).Text) <> Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)) Then
                    
                        Text(ptxTotal_Qty).Locked = False
                    
                    End If
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                                        
                    
        '>>>>>>2018.09.26
                    Text(ptxSoko_No).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Text(ptxRetu).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Text(ptxRen).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Text(ptxDan).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        
        
        '>>>>>>2018.09.26
                    
                    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                        Text(ptxHin_Nai).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                    End If
                Case True
                    
                    Call Clear_field(ptxHin_Name, 1)
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(ptxHin_Gai).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
    
            
            If Right(Combo(pcmbSagyo).Text, 2) = YOIN_FURIKAE Then
                                        '国内外振替時の振替後品番のチェック
                                        '要因マスタ読み込み追加2001.09.18
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Right(Combo(pcmbSagyo).Text, 2), 1))
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Combo(pcmbSagyo).Text, 1))
                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                        Combo(Index).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                        Unload Me
                End Select
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                If Combo(pcmbNaigai).Text = NAIGAI1 Then
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                Else
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                End If
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "入力したコードは登録されていません。（振替え品目）"
                        Text(ptxHin_Gai).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Unload Me
                End Select
    
            End If
            
            sts = List_Disp_Proc()
            Select Case sts
                Case False
                Case True
                    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
                    Else
                        MsgBox "入力したコードには出庫可能な在庫が有りません。"
                        Text(ptxHin_Gai).SetFocus
                        Exit Sub
                    End If
                Case Else
                    Unload Me
            End Select
            
            Select Case Left(Right(Combo(pcmbSagyo).Text, 2), 1)
                Case ACT_ZAITEI_IN
            End Select
            
            
            If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then
            Else
                If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_IDO_IN Then
                
                    List1(PlstZaiko).ListIndex = 0
                    List1(PlstZaiko).SetFocus
                    Exit Sub
                End If
            End If
    
    
        Case ptxTotal_Qty                       '総出庫数のチェック
            If Not IsNumeric(Text(ptxTotal_Qty).Text) Then
                MsgBox "入力した項目はエラーです｡"
                Text(ptxTotal_Qty).SetFocus
                Exit Sub
            End If
    
    
    
    
    
    
            If CLng(Text(ptxTotal_Qty).Text) <= 0 Then
                MsgBox "入力した項目はエラーです｡"
                Text(ptxTotal_Qty).SetFocus
                Exit Sub
            End If
    
            If CLng(Text(ptxTotal_Qty).Text) > CLng(lblTotal(2).Caption) Then
                MsgBox "入力した項目はエラーです｡(総出庫数不足)"
                Text(ptxTotal_Qty).SetFocus
                Exit Sub
            End If
    
    
            Text(ptxTotal_Qty).Locked = True    '以降入力不可
            lbl_ZAN_QTY.Caption = Text(ptxTotal_Qty).Text
    
    
            List1(PlstZaiko).ListIndex = 0
            List1(PlstZaiko).SetFocus
            Exit Sub
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Yoin_Set_Proc() As Integer


'-------------------------------------------
'
'   2007.12.10  表示順を考慮して作業を並べる
'
'-------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim i           As Integer




    Yoin_Set_Proc = True
    
    Combo(pcmbSagyo).Clear
    List2.Clear             '2007.12.10


    For i = 0 To UBound(MENU_TBL)
        If MENU_TBL(i).YOIN = "0" Then
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, MENU_TBL(i).CODE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
                        
            com = BtOpGetGreater
            Do
                sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> MENU_TBL(i).CODE Then
                            Exit Do
                        End If
                    
                        If StrConv(YOINREC.REGI_F, vbUnicode) = "0" Or _
                            StrConv(YOINREC.REGI_F, vbUnicode) = "2" Then
                            
                            '2007.12.10 ↓
                            'Combo(pcmbSagyo).AddItem StrConv(YOINREC.YOIN_DNAME, vbUnicode) & " " & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                            
                            If StrConv(YOINREC.DSP_No, vbUnicode) <> "**" Then
                            
                                List2.AddItem StrConv(YOINREC.DSP_No, vbUnicode) & StrConv(YOINREC.YOIN_DNAME, vbUnicode) & " " & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                            End If
                            '2007.12.10 ↑
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "要因マスタ")
                        Exit Function
                End Select
            
            Loop
        End If
    Next i


    If List2.ListCount = 0 Then
    Else
        For i = 0 To List2.ListCount - 1
        
            Combo(pcmbSagyo).AddItem Mid(List2.List(i), 3, Len(List2.List(i)) - 2)
        Next i
    End If


    If Combo(pcmbSagyo).ListCount = 0 Then
        Exit Function
    End If
    
    Yoin_Set_Proc = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020401)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020401)


    F1020401.MousePointer = vbDefault

End Sub


Private Sub Input_Change_Proc()

    
    
    Select Case Left(Right(Combo(pcmbSagyo).Text, 2), 1)
        Case ACT_ZAITEI_IN                          '入庫
            
            Label(14).Visible = False
            Text(ptxTotal_Qty).Visible = False      '出庫総数   OFF
            lbl_ZAN_T.Visible = False
            lbl_ZAN_QTY.Visible = False
            lbl_ZAN_T.Visible = False
            lbl_ZAN_QTY.Visible = False
            
            Text(ptxSoko_No).TabStop = True         '倉庫番号　 ON
            Text(ptxSoko_No).Locked = False
            Text(ptxRetu).TabStop = True            '列         ON
            Text(ptxRetu).Locked = False
            Text(ptxRen).TabStop = True             '連         ON
            Text(ptxRen).Locked = False
            Text(ptxDan).TabStop = True             '段         ON
            Text(ptxDan).Locked = False
        
            Text(ptxNyuka_DT_YY).TabStop = True     '入荷日     ON
            Text(ptxNyuka_DT_YY).Locked = False
            Text(ptxNyuka_DT_MM).TabStop = True
            Text(ptxNyuka_DT_MM).Locked = False
            Text(ptxNyuka_DT_DD).TabStop = True
            Text(ptxNyuka_DT_DD).Locked = False
        
        
        
            Text(ptxNyuka_DT_YY).Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)      '2017.07.22
            Text(ptxNyuka_DT_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)      '2017.07.22
            Text(ptxNyuka_DT_DD).Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)      '2017.07.22
        
        
        
            List1(PlstZaiko).TabStop = False        'LISTBOX    OFF
        
        
        
            Text(ptxSumi_QTY).TabStop = True
            Text(ptxSumi_QTY).Locked = False
            Text(ptxMi_QTY).TabStop = True
            Text(ptxMi_QTY).Locked = False
        
        
        
        Case ACT_ZAITEI_OUT, ACT_IDO_OUT                '出庫／移動出庫
            
            Label(14).Visible = True
            Text(ptxTotal_Qty).Visible = True       '出庫総数   ON
            Text(ptxTotal_Qty).Text = ""            '入力数クリアー
            Text(ptxTotal_Qty).Locked = False       '入力数クリアー
            lbl_ZAN_T.Visible = True
            lbl_ZAN_QTY.Visible = True
            
            Text(ptxSoko_No).TabStop = False        '倉庫番号   OFF
            Text(ptxSoko_No).Locked = True
            Text(ptxRetu).TabStop = False           '列         OFF
            Text(ptxRetu).Locked = True
            Text(ptxRen).TabStop = False            '連         OFF
            Text(ptxRen).Locked = True
            Text(ptxDan).TabStop = False            '段         OFF
            Text(ptxDan).Locked = True
        
            Text(ptxNyuka_DT_YY).TabStop = False    '入荷日     OFF
            Text(ptxNyuka_DT_YY).Locked = True
            Text(ptxNyuka_DT_MM).TabStop = False
            Text(ptxNyuka_DT_MM).Locked = True
            Text(ptxNyuka_DT_DD).TabStop = False
            Text(ptxNyuka_DT_DD).Locked = True
            
            List1(PlstZaiko).TabStop = True         'LISTBOX    ON
        
        Case ACT_IDO_IN                        '移動入庫
            
            Label(14).Visible = False
            Text(ptxTotal_Qty).Visible = False      '出庫総数   OFF
            lbl_ZAN_T.Visible = False
            lbl_ZAN_QTY.Visible = False
            lbl_ZAN_T.Visible = False
            lbl_ZAN_QTY.Visible = False
            
            Text(ptxSoko_No).TabStop = True         '倉庫番号   ON
            Text(ptxSoko_No).Locked = False
            Text(ptxRetu).TabStop = True            '列         ON
            Text(ptxRetu).Locked = False
            Text(ptxRen).TabStop = True             '連         ON
            Text(ptxRen).Locked = False
            Text(ptxDan).TabStop = True             '段         ON
            Text(ptxDan).Locked = False
        
            Text(ptxNyuka_DT_YY).TabStop = False    '入荷日     OFF
            Text(ptxNyuka_DT_YY).Locked = True
            Text(ptxNyuka_DT_MM).TabStop = False
            Text(ptxNyuka_DT_MM).Locked = True
            Text(ptxNyuka_DT_DD).TabStop = False
            Text(ptxNyuka_DT_DD).Locked = True
            
            Text(ptxSumi_QTY).TabStop = True
            Text(ptxSumi_QTY).Locked = True
            Text(ptxMi_QTY).TabStop = True
            Text(ptxMi_QTY).Locked = True
            
            
            
            List1(PlstZaiko).TabStop = True
    
    End Select

            
'>>>>>>>>>  2017.07.22
    If Text(ptxSoko_No).Locked Then
        Text(ptxSoko_No).BackColor = &H8000000F
    Else
        Text(ptxSoko_No).BackColor = &H80000005
    End If
    
    If Text(ptxRetu).Locked Then
        Text(ptxRetu).BackColor = &H8000000F
    Else
        Text(ptxRetu).BackColor = &H80000005
    End If
    
    If Text(ptxRen).Locked Then
        Text(ptxRen).BackColor = &H8000000F
    Else
        Text(ptxRen).BackColor = &H80000005
    End If
    
    If Text(ptxDan).Locked Then
        Text(ptxDan).BackColor = &H8000000F
    Else
        Text(ptxDan).BackColor = &H80000005
    End If

    If Text(ptxNyuka_DT_YY).Locked Then
        Text(ptxNyuka_DT_YY).BackColor = &H8000000F
    Else
        Text(ptxNyuka_DT_YY).BackColor = &H80000005
    End If
    
    If Text(ptxNyuka_DT_MM).Locked Then
        Text(ptxNyuka_DT_MM).BackColor = &H8000000F
    Else
        Text(ptxNyuka_DT_MM).BackColor = &H80000005
    End If
    
    If Text(ptxNyuka_DT_DD).Locked Then
        Text(ptxNyuka_DT_DD).BackColor = &H8000000F
    Else
        Text(ptxNyuka_DT_DD).BackColor = &H80000005
    End If
    
    If Text(ptxSumi_QTY).Locked Then
        Text(ptxSumi_QTY).BackColor = &H8000000F
    Else
        Text(ptxSumi_QTY).BackColor = &H80000005
    End If
    
    If Text(ptxMi_QTY).Locked Then
        Text(ptxMi_QTY).BackColor = &H8000000F
    Else
        Text(ptxMi_QTY).BackColor = &H80000005
    End If
'>>>>>>>>>  2017.07.22
    



End Sub

Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True

                                        '在庫合計データ使用中チェック
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        Item_Read_Proc = SYS_ERR
        Exit Function
    End If


    If Combo(pcmbNaigai).Text = NAIGAI1 Then            '国内外の選択
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If


    




    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)          '事業部
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)               '国内外
                                                            '品目（外部）
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                                            '内部品番で読み替え
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)  '事業部
            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)       '国内外
                                                            '品目（内部）
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxHin_Gai).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    Text(ptxHin_Gai).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                Case BtErrKeyNotFound
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Item_Read_Proc = SYS_ERR
                    Exit Function
            End Select
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Read_Proc = SYS_ERR
            Exit Function
    End Select

    Item_Read_Proc = False

End Function


Private Sub Clear_field(Optional Start_Field As Integer = 0, Optional Mode As Integer = 0)
Dim i   As Integer

    For i = Start_Field To Text_Max
        Text(i).Text = ""
    Next i


    If Left(Right(Combo(pcmbSagyo).Text, 2), 1) = ACT_ZAITEI_IN Then        '2017.09.23
    
        Text(ptxNyuka_DT_YY).Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)      '2017.09.23
        Text(ptxNyuka_DT_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)      '2017.09.23
        Text(ptxNyuka_DT_DD).Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)      '2017.09.23
    End If                                                                  '2017.09.23



    If Mode = 0 Then
'        lbl_ZAN_T.Visible = False
'        lbl_ZAN_QTY.Visible = False
'        lbl_ZAN_QTY.Caption = ""
        List1(PlstZaiko).Clear

'        lblTotal(0).Caption = ""
'        lblTotal(1).Caption = ""
'        lblTotal(2).Caption = ""
    
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)

'>>>>>>>>>>>>>>>>>>>>>> 小文字→大文字 2016.01.26
'    If Index = ptxHin_Gai Then
'
'        Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
'
'    End If


    Select Case Index
        Case ptxHin_Gai
            Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
        Case ptxSoko_No
            Text(Index).Text = RTrim(StrConv(Text(Index).Text, vbUpperCase))
    End Select
'>>>>>>>>>>>>>>>>>>>>>> 小文字→大文字 2016.01.26
End Sub
