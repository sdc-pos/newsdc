VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "TDBG8.OCX"
Begin VB.Form F1050351 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出庫実績照会"
   ClientHeight    =   10710
   ClientLeft      =   795
   ClientTop       =   -90
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14910
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   8640
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8655
      Left            =   0
      OleObjectBlob   =   "F1050351.frx":0000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14895
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   300
      Index           =   0
      Left            =   1080
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   1
      Top             =   120
      Width           =   2175
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9840
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
      Index           =   8
      Left            =   7800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "再表示"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "棚 札"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "逆 順"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "正 順"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
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
      TabIndex        =   30
      Top             =   10440
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   29
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3480
      TabIndex        =   28
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   2640
      TabIndex        =   27
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   2160
      TabIndex        =   26
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   25
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日付"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番"
      Height          =   252
      Index           =   0
      Left            =   2040
      TabIndex        =   22
      Top             =   240
      Width           =   612
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1050351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0           '国内外

Private Const ptxHin_Gai% = 0           '品番（外部）
Private Const ptxHin_Name% = 1          '品名
Private Const ptxST_DT_YY% = 2          '開始日付 年
Private Const ptxST_DT_MM% = 3          '開始日付 月
Private Const ptxST_DT_DD% = 4          '開始日付 日
Private Const ptxEN_DT_YY% = 5          '終了日付 年
Private Const ptxEN_DT_MM% = 6          '終了日付 月
Private Const ptxEN_DT_DD% = 7          '終了日付 日

Private Const Text_Max% = 7

Dim IDO     As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数
Dim Max_Row     As Long                 'リストボックス最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 22             '最大列数

Private Const ColHin_Gai% = 0           '列 品番（外部）
Private Const ColHin_Name% = 1          '列 品名
Private Const ColRIRK% = 2              '列 履歴
Private Const ColDEN_DT% = 3            '列 伝票日付
Private Const ColBIN_No% = 4            '列 便
Private Const ColDEN_No% = 5            '列 伝票№
Private Const ColNyuko_Qty% = 6         '列 入庫数
Private Const ColSyuko_Qty% = 7         '列 出庫数
Private Const ColZAITEI_Qty% = 8        '列 在庫訂正数
Private Const ColIDO_Qty% = 9           '列 移動数
Private Const ColHin_Zaiko_Qty% = 10    '列 品目別在庫数
Private Const ColMUKE_DNAME% = 11       '列 向け先
Private Const ColTANTO_NAME% = 12       '列 ID
Private Const ColMEMO% = 13             '列 メモ
Private Const ColJITU_DT% = 14          '列 実績日付
Private Const ColJITU_TM% = 15          '列 実績時刻
Private Const ColFrom_Location% = 16    '列 From棚
Private Const ColTO_Location% = 17      '列 To棚
Private Const ColNYUKA_DT% = 18         '列 入荷日
Private Const ColHin_Nai% = 19          '列 品番（内部）
Private Const ColSS_Name% = 20          '列 直送先名
Private Const ColTOKU_MARK% = 21        '列 特売りマーク
Private Const ColID_NO% = 22            '列 伝票ＩＤ

Private Sort_Tbl(ColHin_Gai To ColID_NO) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順

Dim Excel_Put_Flg       As Boolean      '棚札出力有無


Dim Excel_Template      As String       '棚札 ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)
Dim Excel_PutPath       As String       '棚札 書き込み先ﾊﾟｽ

Dim Excel_Put_Yoin_IN   As Variant      '棚札 入庫対象要因配列
Dim Excel_Put_Yoin_OUT  As Variant      '棚札 入庫対象要因配列


Dim Excel_Bin_Mei       As Variant      '棚札 便名称配列
Dim ExcelApp            As Excel.Application
Dim Excelbook           As Excel.Workbook
Dim ExcelWorkSheet      As Excel.Worksheet

Private Function Err_Chk_Proc() As Integer

'日付範囲入力エラーチェック    2007.5.15 (ｻﾌﾞﾌﾟﾛｼｰｼﾞｬ化)

Dim sts         As Integer
Dim ans         As Integer
Dim i           As Integer


    Err_Chk_Proc = True


    For i = ptxST_DT_YY To ptxEN_DT_DD
        Select Case i
            Case ptxST_DT_YY, ptxEN_DT_YY
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_YY Then
                        Text(i).Text = "0000"
                    Else
                        Text(i).Text = "9999"
                    End If
                Else
                    If Not IsNumeric(Text(i).Text) Then
                    Else
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    End If
                End If
            Case Else
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_MM Or i = ptxST_DT_DD Then

                        Text(i).Text = "00"
                    Else
                        Text(i).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(i).Text) Then
                Else
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End Select
    Next i

    If (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) > _
        (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
        MsgBox "入力した項目はエラーです。"
        Text(ptxST_DT_YY).SetFocus
        Exit Function
    End If


    Err_Chk_Proc = False

End Function

Private Function List_Disp_Proc(Mode As Integer) As Integer
                                    '画面表示内容設定
                                    'Mode = 0:昇順
                                    'mode = 1:降順
Dim sts         As Integer
Dim com         As Integer
Dim Key_Mode    As Integer
Dim NAIGAI      As String * 1

Dim ans         As Integer
Dim i           As Integer
Dim Row         As Long

Dim Skip_flg    As Boolean  '2004.07.16

    List_Disp_Proc = True

                                    'エラーチェック
    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("品目マスタは登録されてません。 処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                List_Disp_Proc = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select

    If Err_Chk_Proc Then            '日付範囲入力ｴﾗｰﾁｪｯｸ    2007.5.15 (ｻﾌﾞﾌﾟﾛｼｰｼﾞｬ化)
        List_Disp_Proc = False
        Exit Function
    End If


    Call Input_Lock

                                    'テーブルリセット
    Set IDO = Nothing

    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1                '国内
            NAIGAI = NAIGAI_NAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI2                '海外
            NAIGAI = NAIGAI_GAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI0                '内外指定なし
            Key_Mode = 0
    End Select


                                    '在庫移動歴読み込み開始
    If Key_Mode = 0 Then
                                    '時系列で読み込み
        Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "")           '昇順
        Else
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "zzzzzzzz")   '降順
        End If
                                    '列表示あり 品番／品名
        TDBGrid1.Columns(ColHin_Gai).Visible = True
        TDBGrid1.Columns(ColHin_Name).Visible = True

    Else
                                    '品番＞時系列で読み込む
        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "")           '昇順
        Else
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzzzz")   '降順
        End If
                                    '列表示なし 品番／品名
        TDBGrid1.Columns(ColHin_Gai).Visible = False
        TDBGrid1.Columns(ColHin_Name).Visible = False
    End If


    Row = Min_Row - 1

    If Mode = 0 Then
        com = BtOpGetGreater        '昇順
    Else
        com = BtOpGetLess           '降順
    End If
    Do
        If Key_Mode = 0 Then
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Else
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        End If

        Skip_flg = False

        Select Case sts
            Case BtNoErr

                If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "A" Or _
                    Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "C" Then
                    Skip_flg = True
                End If

            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                List_Disp_Proc = SYS_ERR
        End Select

        If Not Skip_flg Then

                                    '事業部 KEYﾌﾞﾚｰｸ
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '日付範囲外
            If Mode = 0 Then
                                    '昇順
                If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                    Exit Do
                End If
            Else
                                    '降順
                If StrConv(IDOREC.JITU_DT, vbUnicode) < (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) Then
                    Exit Do
                End If
            End If

            If Key_Mode = 1 Then
                                    '品番指定時、品番ﾌﾞﾚｰｸ
                If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
            End If


            If Key_Mode = 0 Then
                If StrConv(IDOREC.NAIGAI, vbUnicode) = NAIGAI Then
                    Row = Row + 1
                    If Row > Max_Row Then
                        Beep
                        MsgBox "最大表示行数を超えました。"
                        Exit Do
                    End If
                    Call Grid_Set_Proc(Row)
                End If
            Else
                Row = Row + 1
                If Row > Max_Row Then
                    Beep
                    MsgBox "最大表示行数を超えました。"
                    Exit Do
                End If

                Call Grid_Set_Proc(Row)
            End If

        End If

        If Mode = 0 Then
            com = BtOpGetNext   '昇順
        Else
            com = BtOpGetPrev   '降順
        End If
        DoEvents
    Loop
                                'DBテーブルリンク
    Set TDBGrid1.Array = IDO
    TDBGrid1.ReBind

    TDBGrid1.Update
    TDBGrid1.MoveFirst


    Call Input_UnLock




    Text(ptxHin_Gai).SetFocus

    List_Disp_Proc = False

End Function

Private Function Tana_Fuda_Put() As Integer

'   部材流通現品棚札　作成                  2007.5.15

Dim strExelFile     As String
Dim Rec_Cnt         As Long
Dim Page_Offset     As Long
Dim posG            As Long

Dim sts             As Integer
Dim com             As Integer
Dim Key_Mode        As Integer
Dim NAIGAI          As String * 1
Dim ans             As Integer
Dim i               As Integer
Dim Skip_flg        As Boolean

'On Error GoTo ERR_PRT


    Tana_Fuda_Put = True

    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        
        
        
        Case NAIGAI0
            MsgBox "国内外は省略できません。", vbExclamation
            Text(ptxHin_Gai).SetFocus
            Tana_Fuda_Put = False
            Exit Function
    
        
    
    End Select
                                    'エラーチェック
    If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
        MsgBox "品番は省略できません。", vbExclamation
        Text(ptxHin_Gai).SetFocus
        Tana_Fuda_Put = False
        Exit Function
    End If

    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("品目マスタは登録されていません。処理を継続しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                Tana_Fuda_Put = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select

    If Err_Chk_Proc Then            '入力ｴﾗｰﾁｪｯｸ    2007.5.15 (ｻﾌﾞﾌﾟﾛｼｰｼﾞｬ化)
        Exit Function
    End If


    Call Input_Lock

                                    '出力ﾌｧｲﾙ名編集
    strExelFile = Excel_PutPath & Trim(Text(ptxHin_Gai).Text) & ".xls"

    'Excelｱﾌﾟﾘｹｰｼｮﾝｵﾌﾞｼﾞｪｸﾄ取得
    Set ExcelApp = CreateObject("Excel.Application")

    Set Excelbook = ExcelApp.Workbooks.Open(Excel_Template)         'ﾃﾝﾌﾟﾚｰﾄﾌﾞｯｸを開く
    Set ExcelWorkSheet = Excelbook.Worksheets(1)                    '１ｼｰﾄ目を選択

    '品番
    ExcelWorkSheet.Application.Cells(3, 2).Value = Trim(Text(ptxHin_Gai).Text)
    '発行日
    ExcelWorkSheet.Application.Cells(1, 8).Value = Format(Now, "yyyy/m/d")

                                    '品番＞時系列で読み込む
    Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)

    Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
    Call UniCode_Conv(K1_IDO.JITU_TM, "")

    Rec_Cnt = 0
    Page_Offset = 6
    posG = 6

    com = BtOpGetGreater
    Do
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)

        Skip_flg = False

        Select Case sts
            Case BtNoErr


            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                Tana_Fuda_Put = SYS_ERR
        End Select

        If Not Skip_flg Then
                                    '事業部 KEYﾌﾞﾚｰｸ
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '日付範囲外
            If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                Exit Do
            End If

                                    '品番指定時、品番ﾌﾞﾚｰｸ
            If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                Exit Do
            End If


            For i = 0 To UBound(Excel_Put_Yoin_IN)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_IN(i) Then
                    Call TanaFuda_Set_Proc(1, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        
            For i = 0 To UBound(Excel_Put_Yoin_OUT)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_OUT(i) Then
                    Call TanaFuda_Set_Proc(2, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        End If

        com = BtOpGetNext
        DoEvents
    Loop

    '当該ページの残り行をクリア
    If posG <= Page_Offset + 35 Then
        Call UniCode_Conv(IDOREC.JITU_DT, "")
        Call UniCode_Conv(IDOREC.BIN_NO, "")
        Call UniCode_Conv(IDOREC.DEN_NO, "")
        Call UniCode_Conv(IDOREC.SUM_KBN, "")
        Call UniCode_Conv(IDOREC.TANTO_NAME, "")
        Call UniCode_Conv(IDOREC.RIRK_NAME, "")
        Call UniCode_Conv(IDOREC.RIRK_ID, "")           '2007.09.06
        Do
            If posG > Page_Offset + 35 Then
                Exit Do
            End If
            Call TanaFuda_Set_Proc(0, posG, Page_Offset)        '１行分編集
        Loop
    End If



    '編集時にﾜｰｸｼｰﾄの先頭が表示される様に「A1」をｱｸﾃｨﾌﾞにする
    ExcelWorkSheet.Application.Range("A1").Activate

    ExcelApp.DisplayAlerts = False              'ﾏｸﾛ実行ｴﾗｰは表示しない


    If Rec_Cnt > 0 Then
        On Error Resume Next
     '   Kill strExelFile
        ExcelWorkSheet.SaveAs strExelFile
'        On Error GoTo 0
    End If


    ExcelApp.Visible = False
    ExcelApp.Workbooks.Close                                        'ﾜｰｸﾌﾞｯｸ閉じる
    ExcelApp.Quit

    Set ExcelWorkSheet = Nothing                                    'ﾜｰｸｼｰﾄ開放
    Set Excelbook = Nothing                                         'ﾜｰｸﾌﾞｯｸ開放

    Set ExcelApp = Nothing                                         'ﾜｰｸﾌﾞｯｸ開放


    Call Input_UnLock

    Text(ptxHin_Gai).SetFocus


    Tana_Fuda_Put = False

End Function

Private Sub TanaFuda_Set_Proc(InOut As Integer, posG As Long, Page_Offset As Long)


'InOut =0(DUMMY) =1(In) =2()


Dim c   As String * 128


    '１頁分編集完了⇒次頁分のフォーマットをコピー
    If posG > Page_Offset + 35 Then
        ExcelWorkSheet.Application.Range(Page_Offset & ":" & Page_Offset + 35).Copy
        ExcelWorkSheet.Application.Range(Page_Offset + 36 & ":" & Page_Offset + 71).Select
        ExcelWorkSheet.Paste

        Page_Offset = Page_Offset + 36
        posG = Page_Offset
    End If

                                            '実績日付
    If Len(Trim(StrConv(IDOREC.JITU_DT, vbUnicode))) <> 0 Then
        ExcelWorkSheet.Application.Cells(posG, 1).Value = _
                                          Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
    Else
        ExcelWorkSheet.Application.Cells(posG, 1).Value = ""
    End If
                                            '便
    Select Case StrConv(IDOREC.BIN_NO, vbUnicode)
        Case "01"         '１便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(0)
        Case "02"         '２便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(1)
        Case "03"         '３便
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(2)
        Case Else
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Trim(StrConv(IDOREC.BIN_NO, vbUnicode))
    End Select
                                            '伝票№
    If InOut = 1 Then
        ExcelWorkSheet.Application.Cells(posG, 3).Value = Trim(StrConv(IDOREC.DEN_NO, vbUnicode))
    Else
        ExcelWorkSheet.Application.Cells(posG, 3).Value = ""    '2007.09.06
    End If
                                            '実績数
    ExcelWorkSheet.Application.Cells(posG, 4).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 5).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 6).Value = ""
    Select Case InOut
        Case 1         '入庫数
            ExcelWorkSheet.Application.Cells(posG, 4).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        
        
                                            '品目別在庫数
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
        
        Case 2         '出庫数
            ExcelWorkSheet.Application.Cells(posG, 5).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
    
                                            '品目別在庫数
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
    
    End Select
                                            '担当者
    ExcelWorkSheet.Application.Cells(posG, 7).Value = Trim(StrConv(IDOREC.TANTO_NAME, vbUnicode))
                                        
                                        '履歴（要因名称）
    If GetIni(App.EXEName, StrConv(IDOREC.RIRK_ID, vbUnicode), "SYS", c) Then
        ExcelWorkSheet.Application.Cells(posG, 8).Value = ""
    Else
        ExcelWorkSheet.Application.Cells(posG, 8).Value = Trim(c)
    End If
    
    

    posG = posG + 1

End Sub
                                    '画面初期状態を設定する
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer
   
    For i = Mode To Text_Max
        Text(i).Text = ""
    Next i
    
End Sub
                                    '品目マスタより各項目を表示する
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True
                                                '国内外の判定
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        Case NAIGAI0
            Text(ptxHin_Gai).Text = ""
            Text(ptxHin_Name).Text = ""
            Item_Read_Proc = False
            Exit Function
    End Select
                                                
    If Len(Text(ptxHin_Gai).Text) = 0 Then
        Text(ptxHin_Name).Text = ""
        Item_Read_Proc = False
        Exit Function
    End If
                                                'まず外部品番で読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '内部品番で再度読み込み
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxHin_Gai).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    
                    Text(ptxHin_Gai).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    
        
                Case BtErrKeyNotFound
        
                    Exit Function
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Item_Read_Proc = SYS_ERR
            End Select
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Read_Proc = SYS_ERR
    End Select
            
    Item_Read_Proc = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbNAIGAI
            Call Clear_Field(0)
            
            If Combo(Index).Text = NAIGAI0 Then
                Text(ptxHin_Gai).Text = ""
                Text(ptxHin_Name).Text = ""
                Text(ptxST_DT_YY).SetFocus
            Else
                Text(ptxHin_Gai).SetFocus
            End If
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim sts As Integer
Dim i   As Integer


On Error Resume Next
    Select Case Index
        Case 0                           '正順表示
            If List_Disp_Proc(0) Then
                Unload Me
            End If

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
            Next i


'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_ASCEND, XTYPE_DATE, 5, XORDER_ASCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub
        Case 3                              '逆順表示
            If List_Disp_Proc(1) Then
                Unload Me
            End If
'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_DESCEND, XTYPE_DATE, 5, XORDER_DESCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
            Next i
            Sort_Tbl(ColJITU_DT) = 1        'ﾃﾞﾌｫﾙﾄ降順
            Sort_Tbl(ColJITU_TM) = 1        'ﾃﾞﾌｫﾙﾄ降順


        Case 4                             '棚札        2007.5.15
            If Tana_Fuda_Put() Then
                Unload Me
            End If


        Case 7                             '再表示
            If List_Disp_Proc(0) Then
                Unload Me
            End If

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
            Next i


        Case 8                             'ﾃﾞｰﾀ出力

            Call Select_Set_Proc

            F1050352.Show vbModal

        Case 11                            '終了
            Unload Me
        Case Else
            Beep
    End Select

End Sub

Private Sub Form_DblClick()
    
Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
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
    LOG_F = RTrim(c)
                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CLng(RTrim(c))








    '棚札用定義情報取り込み 2007.05.15      ここから

                                            
                                            
                    '棚札出力有無
    If GetIni(App.EXEName, "Excel_Put", "SYS", c) Then
        Excel_Put_Flg = False
    Else
        If Trim(c) = "1" Then
            Excel_Put_Flg = True
        Else
            Excel_Put_Flg = False
        End If
    End If
                                            
                                            
    If Excel_Put_Flg Then
                                                'ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)
        If GetIni(App.EXEName, "F105035_EXCEL_TEMPLATE", "SYS", c) Then
            Beep
            MsgBox "ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Template = Trim(c)
                                                '書き込み先ﾊﾟｽ
        If GetIni(App.EXEName, "F105035_EXCEL_OUTPUT", "SYS", c) Then
            Beep
            MsgBox "書き込み先ﾊﾟｽの獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_PutPath = Trim(c)
                                                '対象入庫要因配列
        If GetIni(App.EXEName, "YOIN_IN", "SYS", c) Then
            Beep
            MsgBox "対象入庫要因配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Put_Yoin_IN = Split(Trim(c), ",", -1)
                                                '対象出庫要因配列
        If GetIni(App.EXEName, "YOIN_OUT", "SYS", c) Then
            Beep
            MsgBox "対象出庫要因配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Put_Yoin_OUT = Split(Trim(c), ",", -1)
                                                '便名称配列
        If GetIni("F105035", "BIN", "SYS", c) Then
            Beep
            MsgBox "便名称配列の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
        Excel_Bin_Mei = Split(Trim(c), ",", -1)
    
    
    Else
    
        Command(4).Enabled = False
    End If
    '棚札用定義情報取り込み 2007.05.15      ここまで



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
            F1050351.Caption = "入出庫実績照会（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '国内外取り込み
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).AddItem NAIGAI0
    Combo(pcmbNAIGAI).Text = NAIGAI1
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If


    Load F1050352
    Load SDC_FLD_F
                                '画面初期設定
    Call Clear_Field(0)


    TDBGrid1.Columns(ColHin_Gai).Visible = False
    TDBGrid1.Columns(ColHin_Name).Visible = False

    TDBGrid1.Style.Locked = True

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
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1050351 = Nothing
    Set F1050352 = Nothing
    Set SDC_FLD_F = Nothing

    End
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
    F1050351.Caption = "入出庫実績照会（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        IDO.QuickSort Min_Row, IDO.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = IDO
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


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

Dim i   As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxHin_Gai             '品番
            
            If (Combo(pcmbNAIGAI).Text = NAIGAI0 Or _
                Len(Trim(Text(ptxHin_Gai).Text)) = 0) Then
            Else
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Text(ptxHin_Name).Text = ""
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
                        
        Case ptxST_DT_YY, ptxEN_DT_YY
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_YY Then
                    Text(Index).Text = "0000"
                Else
                    Text(Index).Text = "9999"
                End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "0000")
                End If
            End If
    
        Case ptxST_DT_MM, ptxST_DT_DD, ptxEN_DT_MM, ptxEN_DT_DD
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_MM Or Index = ptxST_DT_DD Then
                                
                    Text(Index).Text = "00"
                Else
                    Text(Index).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
    
            If Index = ptxEN_DT_DD Then
                If List_Disp_Proc(0) Then
                    Unload Me
                End If
            End If
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1050351.MousePointer = vbHourglass

    Call Ctrl_Lock(F1050351)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1050351)


    F1050351.MousePointer = vbDefault

End Sub

Private Sub Grid_Set_Proc(Row As Long)


    IDO.ReDim Min_Row, Row, Min_Col, Max_Col
                                            '品目（外部）
    IDO(Row, ColHin_Gai) = StrConv(IDOREC.HIN_GAI, vbUnicode)       '品目（外部）
                                            '品名
    IDO(Row, ColHin_Name) = StrConv(IDOREC.HIN_NAME, vbUnicode)     '品目名称
                                            '履歴（要因）
    IDO(Row, ColRIRK) = StrConv(IDOREC.RIRK_NAME, vbUnicode)        '要因名称
                                            '特売りマーク
    IDO(Row, ColTOKU_MARK) = StrConv(IDOREC.TOKU_MARK, vbUnicode)   '特売りマーク
                                            '実績日付
    IDO(Row, ColJITU_DT) = Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
                                            '実績時刻
    IDO(Row, ColJITU_TM) = Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 1, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 3, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 5, 2)
                                            '伝票日付
    If Len(Trim(StrConv(IDOREC.DEN_DT, vbUnicode))) <> 0 Then
        IDO(Row, ColDEN_DT) = Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 7, 2)
    End If
                                            '便         2007.5.16
    IDO(Row, ColBIN_No) = StrConv(IDOREC.BIN_NO, vbUnicode)
                                            '伝票№
    IDO(Row, ColDEN_No) = StrConv(IDOREC.DEN_NO, vbUnicode)
                                            '品目別在庫数
    IDO(Row, ColHin_Zaiko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + CLng(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode)), "#,##0")
                                            '実績数
    Select Case StrConv(IDOREC.SUM_KBN, vbUnicode)
        Case SUM_KBN_IN
                                            '入庫数
            IDO(Row, ColNyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
        Case SUM_KBN_OT
                                            '出庫数
            IDO(Row, ColSyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                
        Case SUM_KBN_ZT
                If Mid(StrConv(IDOREC.RIRK_ID, vbUnicode), 1, 1) = ACT_ZAITEI_IN Then
                                            '在訂（＋）
                    IDO(Row, ColZAITEI_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                Else
                                            '在訂（－）
                    IDO(Row, ColZAITEI_Qty) = Format((CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))) * -1), "#,##0")
                End If
        
        Case SUM_KBN_MV
                                            '移動数
                IDO(Row, ColIDO_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
    End Select
                                            'FROM棚
    If Len(Trim(StrConv(IDOREC.FROM_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColFrom_Location) = StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_DAN, vbUnicode)
    End If
                                            'TO棚
    If Len(Trim(StrConv(IDOREC.TO_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColTO_Location) = StrConv(IDOREC.TO_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_DAN, vbUnicode)
    End If
                                            '入荷日
    If Len(Trim(StrConv(IDOREC.NYUKA_DT, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColNYUKA_DT) = Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 7, 2)
    End If
                                            '向け先
    IDO(Row, ColMUKE_DNAME) = StrConv(IDOREC.MUKE_DNAME, vbUnicode)
                                            '担当者
    IDO(Row, ColTANTO_NAME) = StrConv(IDOREC.TANTO_NAME, vbUnicode)
                                            '品番（内部）
    IDO(Row, ColHin_Nai) = StrConv(IDOREC.HIN_NAI, vbUnicode)
                                            'メモ
    IDO(Row, ColMEMO) = StrConv(IDOREC.MEMO, vbUnicode)
                                            '伝票ＩＤ
    IDO(Row, ColID_NO) = StrConv(IDOREC.ID_NO, vbUnicode)
                            
'    TDBGrid1.Update
End Sub

Private Sub Select_Set_Proc()

    F1050352.Combo(pcmbNAIGAI).Text = Combo(pcmbNAIGAI).Text
    F1050352.Text(ptxHin_Gai).Text = Text(ptxHin_Gai).Text
    F1050352.Text(ptxHin_Name).Text = Text(ptxHin_Name).Text
    F1050352.Text(ptxST_DT_YY).Text = Text(ptxST_DT_YY).Text
    F1050352.Text(ptxST_DT_MM).Text = Text(ptxST_DT_MM).Text
    F1050352.Text(ptxST_DT_DD).Text = Text(ptxST_DT_DD).Text
    If Len(Trim(Text(ptxEN_DT_YY).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_YY).Text = "9999"
    Else
        F1050352.Text(ptxEN_DT_YY).Text = Text(ptxEN_DT_YY).Text
    End If
        
    If Len(Trim(Text(ptxEN_DT_MM).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_MM).Text = "99"
    Else
        F1050352.Text(ptxEN_DT_MM).Text = Text(ptxEN_DT_MM).Text
    End If
    
    If Len(Trim(Text(ptxEN_DT_DD).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_DD).Text = "99"
    Else
        F1050352.Text(ptxEN_DT_DD).Text = Text(ptxEN_DT_DD).Text
    End If
End Sub
